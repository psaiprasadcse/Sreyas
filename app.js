/* GitHub Pages single-page dashboard: all Excel processing stays in the browser. */
(function () {
  'use strict';
  const C=window.ScopusDashboardCore;
  const $=id=>document.getElementById(id);
  const state={data:null,selectedAuthor:null,year:null,tab:'overview',page:1,loadedNames:[],demo:false,published:false};
  let toastTimer;
  const palette={Q1:'#178c85',Q2:'#597bd4',Q3:'#e6a64e',Q4:'#cf7087','N/A':'#b7bfcd'};
  const formatN=v=>Number(v||0).toLocaleString('en-IN');
  const text=(el,content)=>{el.textContent=content==null?'':String(content);return el;};
  function toast(message){const t=$('toast');text(t,message);t.classList.add('visible');clearTimeout(toastTimer);toastTimer=setTimeout(()=>t.classList.remove('visible'),5100);}
  function safeUrl(value){try{const u=new URL(String(value||''));return ['https:','http:'].includes(u.protocol)?u.toString():'';}catch(_){return '';}}
  function makeLink(label,href){const link=document.createElement('a'); const url=safeUrl(href); if(!url) return text(document.createElement('span'),'—');link.href=url;link.target='_blank';link.rel='noopener noreferrer';text(link,label);return link;}
  function makeCell(tr,contents,tag='td'){const td=document.createElement(tag);if(contents instanceof Node) td.appendChild(contents);else text(td,contents);tr.appendChild(td);return td;}
  function badge(q){const el=document.createElement('span');el.className='quartile-tag '+(q==='N/A'?'':q);text(el,q);return el;}
  function csvEscape(v){const raw=String(v==null?'':v);const safe=/^[=+@\-\t\r]/.test(raw)?"'"+raw:raw;return '"'+safe.replace(/"/g,'""')+'"';}
  function downloadCsv(filename,headers,rows){const csv='\ufeff'+[headers,...rows].map(r=>r.map(csvEscape).join(',')).join('\r\n');const blob=new Blob([csv],{type:'text/csv;charset=utf-8'});const url=URL.createObjectURL(blob);const link=document.createElement('a');link.href=url;link.download=filename;document.body.appendChild(link);link.click();link.remove();setTimeout(()=>URL.revokeObjectURL(url),1500);toast('CSV downloaded: '+filename);}
  function datasetYear(){return state.year||'';}
  function currentPapers(){return C.matchingRecords(state.data,state.selectedAuthor,state.year);}
  function authorByKey(key){return state.data.authors.find(a=>a.key===key);}
  function nameInitials(s){return String(s||'').split(/[\s,]+/).filter(Boolean).slice(0,2).map(t=>t[0].toUpperCase()).join('');}
  function filteredRank(){return C.rankAuthors(state.data,state.year);}
  function loadReports(reports,demo=false,published=false){
    try{
      const data=C.ingestReports(reports);
      if(!data.authors.length) throw Error('No author records found. Please upload a full report with the Monthly Summary and Scopus Records worksheets.');
      state.data=data;state.year=data.years[0]||null;state.selectedAuthor=null;state.page=1;state.demo=demo;state.loadedNames=reports.map(x=>x.name);state.published=published;state.tab='overview';
      const options=$('yearSelect');options.replaceChildren();const all=new Option('All years','');options.add(all);for(const y of data.years)options.add(new Option(String(y),String(y)));options.value=String(state.year||'');
      $('welcomePanel').hidden=true;$('dashboard').hidden=false;$('authorSearch').value='';$('rankingSearch').value='';$('paperSearch').value='';$('journalSearch').value='';
      document.querySelectorAll('.tab').forEach(x=>{x.classList.toggle('active',x.dataset.tab==='overview');x.setAttribute('aria-selected',String(x.dataset.tab==='overview'));});
      document.querySelectorAll('.view').forEach(x=>x.hidden=x.id!=='view-overview');
      renderAll();toast(`Loaded ${data.authors.length} authors and ${formatN(data.records.length)} author–paper records.`);
    } catch(err) {toast('Could not load reports: '+(err.message||err)); console.error(err);}
  }
  function workbookToReport(wb,name){const sheets={};for(const label of wb.SheetNames){const canonical=['Monthly Summary','Scopus Records','Journal Metrics','Run Info','Errors'].find(n=>n.toLowerCase()===label.trim().toLowerCase());if(canonical)sheets[canonical]=XLSX.utils.sheet_to_json(wb.Sheets[label],{defval:'',raw:false});}return (sheets['Monthly Summary']||sheets['Scopus Records'])?{name,sheets}:null;}
  async function parseFiles(fileList){const list=Array.from(fileList||[]);if(!list.length)return;
    if(!window.XLSX){toast('The Excel reader did not load. Check your internet connection and reload.');return;}
    const reports=[];
    for(const file of list){try{
      const bytes=await file.arrayBuffer();const wb=XLSX.read(bytes,{type:'array',cellDates:false});
      const report=workbookToReport(wb,file.name);
      if(!report){toast(`${file.name} is not a Scopus monthly report; skipped.`);continue;}
      reports.push(report);
    }catch(err){toast(`Cannot open ${file.name}: ${err.message||err}`);}}
    if(reports.length)loadReports(reports,false);else toast('No compatible Scopus report was found.');
  }
  async function loadOptionalPublishedReport(){
    // Opt-in public dataset: only when site operator deliberately commits data/report.xlsx.
    if(location.protocol==='file:'||!window.XLSX||state.data) return;
    try {const response=await fetch('./data/report.xlsx',{cache:'no-store'});if(!response.ok||state.data)return;
      const bytes=await response.arrayBuffer();const wb=XLSX.read(bytes,{type:'array',cellDates:false});
      const report=workbookToReport(wb,'Published institutional report');
      if(report&&!state.data)loadReports([report],false,true);
    } catch(e) {console.info('No optional published data file available. Upload an Excel report to begin.');}
  }
  function demoReports(){
    const authors=[
      ['Dr Asha Sharma','11111111111'],['Dr Ravi Kumar','22222222222'],['Dr Meera Joshi','33333333333'],['Dr Arjun Patil','44444444444']
    ];
    const mock=[
      [0,'Explainable AI for Clinical Imaging','2026-01-11','Computational Medical Research','Q1',9.4,'10.0000/demo1','2025'],
      [0,'IoT Monitoring in Smart Cities','2026-03-15','Applied Engineering Systems','Q2',5.8,'10.0000/demo2','2025'],
      [0,'Graph Learning for Diagnostics','2026-05-25','Computational Medical Research','Q1',9.4,'10.0000/demo3','2025'],
      [1,'Energy Forecasting with Machine Learning','2026-02-08','Energy Informatics Review','Q1',8.1,'10.0000/demo4','2025'],
      [1,'Reliable Wireless Sensor Networks','2026-07-19','Applied Engineering Systems','Q2',5.8,'10.0000/demo5','2025'],
      [2,'Battery Health Estimation with Transformers','2026-04-22','Renewable Data Studies','Q3',3.1,'10.0000/demo6','2025'],
      [2,'Low-Cost Environmental Sensors','2026-09-10','Engineering Field Reports','Q4',1.1,'10.0000/demo7','2025'],
      [3,'Anomaly Detection in Manufacturing','2026-06-14','Industrial AI Reports','N/A','', '10.0000/demo8','2025'],
      [3,'Vision-Based Road Safety Assessment','2026-10-29','Applied Engineering Systems','Q2',5.8,'10.0000/demo9','2025'],
      [0,'Graph Learning for Diagnostics','2026-05-25','Computational Medical Research','Q1',9.4,'10.0000/demo3','2025']
    ];
    const rows=mock.map(([a,title,date,journal,q,score,doi,my],i)=>({'Author Name':authors[a][0],'Scopus Author ID':authors[a][1], 'Scopus Profile URL':`https://www.scopus.com/authid/detail.uri?authorId=${authors[a][1]}`,
      'Paper Title':title,'DOI':doi,'Publication Date':date,'Publication Month':C.MONTHS[+date.substring(5,7)-1],'Scopus Search Year':2026,
      'Journal / Source':journal,'Scopus Paper URL':'','Unique Paper Key':'doi:'+doi,'CiteScore':score,'CiteScore Year':my,'Best CiteScore Quartile':q,
      'Highest CiteScore Percentile':{Q1:93,Q2:68,Q3:40,Q4:17,'N/A':''}[q],'Best Subject Category':'Computer Science','Journal Metrics Status':q==='N/A'?'Metric not available':'Complete'}));
    const sm=authors.map(([name,id])=>({'Author Name':name,'Scopus Author ID':id,'Scopus Profile URL':`https://www.scopus.com/authid/detail.uri?authorId=${id}`,'Retrieval Status':'Complete','Annual Total':mock.filter(x=>authors[x[0]][1]===id).length}));
    return [{name:'DEMONSTRATION — fictional data.xlsx',sheets:{'Monthly Summary':sm,'Scopus Records':rows,'Run Info':[{'Report parameter':'Selected year',Value:2026}]}}];
  }
  function selectAuthor(key){state.selectedAuthor=key||null;state.page=1;const author=authorByKey(state.selectedAuthor);$('authorSearch').value=author?author.name:'';renderAll();}
  function renderDirectory(){const holder=$('authorDirectory');holder.replaceChildren();if(!state.data)return;
    const q=C.norm($('authorSearch').value);const filtered=state.data.authors.filter(a=>!q||C.norm(a.name).includes(q)||a.id.includes(q)).slice(0,24);
    const button=(label,key,subtitle)=>{const b=document.createElement('button');b.type='button';b.className='author-option'+(state.selectedAuthor===key?' active':'');const av=document.createElement('span');av.className='author-avatar';text(av,key?nameInitials(label):'★');b.appendChild(av);const inside=document.createElement('span');text(inside,label);b.appendChild(inside);b.addEventListener('click',()=>selectAuthor(key));holder.appendChild(b);};
    button('All authors',null);
    filtered.forEach(a=>button(a.name,a.key));if(q&&filtered.length===0){const msg=document.createElement('div');msg.className='directory-info';text(msg,'No matching authors');holder.appendChild(msg);}
    const n=state.data.authors.filter(a=>!q||C.norm(a.name).includes(q)||a.id.includes(q)).length;
    if(n>24){const msg=document.createElement('div');msg.className='directory-info';text(msg,`Showing 24 of ${formatN(n)} — refine your search`);holder.appendChild(msg);}
  }
  function renderContext(){const data=state.data,selected=authorByKey(state.selectedAuthor),records=currentPapers();
    text($('contextEyebrow'),selected?'AUTHOR PROFILE':'INSTITUTION OVERVIEW');text($('contextTitle'),selected?selected.name:'All authors');
    text($('contextSubtitle'),selected?'Personal Scopus publication statistics and journal metrics':`${formatN(data.authors.length)} Scopus author profiles · ${formatN(records.length)} unique papers`);
    const links=$('profileLinks');links.replaceChildren();if(selected){if(selected.profile)links.append(makeLink('Open Scopus profile ↗',selected.profile));if(selected.scholar)links.append(makeLink('Google Scholar ↗',selected.scholar));}
    text($('dataBadge'),state.demo?'DEMO DATA':'EXCEL LOADED');text($('monthlyYear'),state.year?String(state.year):'ALL YEARS');text($('fileStatus'),'Source: '+state.loadedNames.join('; '));
    const notices=[];
    if(state.demo) notices.push('Demonstration data only. All names and publication records are fictional.');
    if(state.published) notices.push('This report was provided as a publicly accessible GitHub Pages dataset by the site operator.');
    if(data.missingRecordSheets>0) notices.push(`Warning: ${data.missingRecordSheets} report(s) have no detailed Scopus Records. Counts may be incomplete.`);
    if(data.errors.length)notices.push(`${data.errors.length} retrieval error(s) are recorded in the imported workbook. Review the original Errors worksheet before using these results for evaluation.`);
    if(selected&&selected.status&&!/complete|success|ok|retrieved/i.test(selected.status))notices.push('Author retrieval status: '+selected.status);
    if(selected&&state.year&&selected.reportedTotals.has(state.year)&&selected.reportedTotals.get(state.year)!==records.length)notices.push(`Source Summary reports ${selected.reportedTotals.get(state.year)} papers, but ${records.length} distinct author records are available. Check for missing or duplicated data.`);
    const notice=$('dataNotice');notice.hidden=notices.length===0;text(notice,notices.join(' '));
  }
  function renderOverview(){const rec=currentPapers(),s=C.summarize(rec);text($('statTotal'),formatN(s.total));text($('statQ1'),formatN(s.quartiles.Q1));text($('statQ2'),formatN(s.quartiles.Q2));text($('statQ3'),formatN(s.quartiles.Q3));text($('statQ4'),formatN(s.quartiles.Q4));text($('statNA'),formatN(s.quartiles['N/A']));text($('statTotalNote'),state.selectedAuthor?'Author publication records':'Unique papers (co-authorship deduplicated)');
    const chart=$('monthChart');chart.replaceChildren();const vals=C.MONTHS.map(m=>s.months[m]),max=Math.max(1,...vals);
    C.MONTHS.forEach((m,i)=>{const col=document.createElement('div');col.className='month-col';const bar=document.createElement('div');bar.className='month-bar'+(vals[i]===Math.max(...vals)&&vals[i]>0?' highlight':'');bar.style.height=Math.max(2,Math.round(vals[i]/max*93))+'%';bar.title=m+': '+vals[i]+' paper(s)';const val=document.createElement('span');val.className='month-val';text(val,vals[i]);bar.appendChild(val);const lab=document.createElement('span');lab.className='month-label';text(lab,m.slice(0,3));col.append(bar,lab);chart.appendChild(col);});
    const donut=$('quartileDonut'),legend=$('quartileLegend');legend.replaceChildren();text($('donutCenter'),formatN(s.total));let current=0;const slices=[];
    C.QUARTILES.forEach(q=>{const n=s.quartiles[q];const pct=s.total?n/s.total*100:0;const next=current+pct;if(n)slices.push(`${palette[q]} ${current.toFixed(2)}% ${next.toFixed(2)}%`);current=next;const row=document.createElement('div');row.className='legend-line';const dot=document.createElement('i');dot.className='dot '+(q==='N/A'?'qna':q.toLowerCase());const title=document.createElement('span');text(title,q);const num=document.createElement('strong');text(num,formatN(n));row.append(dot,title,num);legend.appendChild(row);});donut.style.background=slices.length?'conic-gradient('+slices.join(',')+')':'#e1e8f0';
    const counts=new Map();for(const p of rec)if(p.journal){const key=C.norm(p.journal);const obj=counts.get(key)||{name:p.journal,count:0};obj.count++;counts.set(key,obj);}const top=Array.from(counts.values()).sort((a,b)=>b.count-a.count||a.name.localeCompare(b.name)).slice(0,6);
    const host=$('topJournals');host.replaceChildren();if(!top.length){const el=document.createElement('div');el.className='empty-state';text(el,'No publication records for this selection.');host.appendChild(el);}for(const x of top){const line=document.createElement('div');line.className='compact-item';const name=document.createElement('span');name.className='item-name';name.title=x.name;text(name,x.name);const count=document.createElement('span');count.className='item-count';text(count,x.count+' papers');line.append(name,count);host.appendChild(line);}
    const facts=[['Publishing sources',formatN(s.journals)],['Classified Q1–Q4',formatN(s.classified)+' of '+formatN(s.total)],['Highest quartile share',s.classified?(s.quartiles.Q1/s.classified*100).toFixed(1)+'% Q1':'N/A'],['Mean journal CiteScore',s.meanCiteScore===null?'N/A':s.meanCiteScore.toFixed(2)],['Unknown publication month',formatN(s.months.Unknown)],['Selected year',state.year||'All years']];
    const container=$('atGlance');container.replaceChildren();for(const [title,val] of facts){const row=document.createElement('div');row.className='fact-row';const a=document.createElement('span');const b=document.createElement('strong');text(a,title);text(b,val);row.append(a,b);container.appendChild(row);}
  }
  function filteredPapers(){let rows=currentPapers();const q=C.norm($('paperSearch').value), quartile=$('paperQuartile').value,journal=$('paperJournal').value,sort=$('paperSort').value;
    rows=rows.filter(p=>(!q||[p.title,p.journal,p.doi,p.authorName].some(v=>C.norm(v).includes(q)))&&(!quartile||p.quartile===quartile)&&(!journal||C.norm(p.journal)===journal));
    rows.sort((a,b)=>{if(sort==='citescore')return (b.citeScore??-1)-(a.citeScore??-1)||String(b.date).localeCompare(String(a.date));if(sort==='title')return a.title.localeCompare(b.title);if(sort==='oldest')return String(a.date).localeCompare(String(b.date));return String(b.date).localeCompare(String(a.date));});return rows;
  }
  function updatePaperJournalOptions(){const sel=$('paperJournal'),prev=sel.value;const names=Array.from(new Set(currentPapers().map(x=>x.journal).filter(Boolean))).sort();sel.replaceChildren(new Option('All journals',''));for(const n of names)sel.add(new Option(n,C.norm(n)));sel.value=Array.from(sel.options).some(o=>o.value===prev)?prev:'';}
  function renderPapers(){const rows=filteredPapers(),size=25,pages=Math.max(1,Math.ceil(rows.length/size));state.page=Math.max(1,Math.min(state.page,pages));const visible=rows.slice((state.page-1)*size,state.page*size);
    text($('paperCount'),`Showing ${formatN(visible.length)} of ${formatN(rows.length)} papers`);text($('pageInfo'),state.page+' / '+pages);$('pagePrev').disabled=state.page<=1;$('pageNext').disabled=state.page>=pages;
    const tbody=$('paperRows');tbody.replaceChildren();if(!visible.length){const tr=document.createElement('tr'),td=makeCell(tr,'No publications match these filters.');td.colSpan=7;td.className='empty-state';tbody.appendChild(tr);return;}
    for(const r of visible){const tr=document.createElement('tr');const paper=makeCell(tr,'');const strong=document.createElement('strong');text(strong,r.title);paper.appendChild(strong);const sub=document.createElement('span');sub.className='subline';text(sub,r.doi?'DOI: '+r.doi:r.eid?'Scopus EID: '+r.eid:'No DOI recorded');paper.appendChild(sub);
      makeCell(tr,state.selectedAuthor?'—':r.authorName);makeCell(tr,r.date||'Unknown');makeCell(tr,r.journal||'Unknown');makeCell(tr,badge(r.quartile));makeCell(tr,r.citeScore===null?'N/A':r.citeScore.toFixed(2));const l=makeCell(tr,'');const links=document.createElement('div');links.className='link-list';if(r.scopusUrl)links.appendChild(makeLink('Scopus ↗',r.scopusUrl));if(r.doiUrl)links.appendChild(makeLink('DOI ↗',r.doiUrl));if(!links.children.length)text(links,'—');l.appendChild(links);tbody.appendChild(tr);}
  }
  function renderAuthors(){const q=C.norm($('rankingSearch').value);const ranking=filteredRank().filter(x=>!q||C.norm(x.author.name).includes(q)||x.author.id.includes(q));text($('rankingCount'),`${formatN(ranking.length)} author(s)`);const tbody=$('authorRows');tbody.replaceChildren();if(!ranking.length){const tr=document.createElement('tr'),td=makeCell(tr,'No matching authors.');td.colSpan=9;td.className='empty-state';tbody.appendChild(tr);return;}
    ranking.forEach((r,i)=>{const tr=document.createElement('tr');makeCell(tr,i+1);const td=makeCell(tr,'');const name=document.createElement('button');name.type='button';name.className='rank-name';text(name,r.author.name);name.addEventListener('click',()=>{selectAuthor(r.author.key);switchTab('overview');window.scrollTo({top:0,behavior:'smooth'});});td.appendChild(name);makeCell(tr,r.total);for(const q of C.QUARTILES)makeCell(tr,r.quartiles[q]);makeCell(tr,r.author.profile?makeLink('View ↗',r.author.profile):'—');tbody.appendChild(tr);});
  }
  function journalStatistics(){const rows=currentPapers(),map=new Map();for(const r of rows){const key=r.sourceId||C.norm(r.journal||'Unknown'),m=map.get(key)||{journal:r.journal||'Unknown',sourceId:r.sourceId,issn:r.issn,count:0,score:null,year:null,pct:null,quartile:'N/A',category:'',link:r.sourceUrl};m.count++;if(r.citeScore!==null&&(m.score===null||(r.citeScoreYear||0)>=(m.year||0))){m.score=r.citeScore;m.year=r.citeScoreYear;m.pct=r.percentile;m.quartile=r.quartile;m.category=r.category;m.link=r.sourceUrl||m.link;}else if(m.quartile==='N/A'&&r.quartile!=='N/A'){m.quartile=r.quartile;m.category=r.category;m.pct=r.percentile;}map.set(key,m);}return Array.from(map.values()).sort((a,b)=>b.count-a.count||a.journal.localeCompare(b.journal));}
  function filteredJournals(){const q=C.norm($('journalSearch').value),quartile=$('journalQuartile').value;return journalStatistics().filter(j=>(!q||[j.journal,j.category,j.issn].some(x=>C.norm(x).includes(q)))&&(!quartile||j.quartile===quartile));}
  function renderJournals(){const rows=filteredJournals(),tbody=$('journalRows');tbody.replaceChildren();text($('journalCount'),`${formatN(rows.length)} sources shown`);if(!rows.length){const tr=document.createElement('tr');const td=makeCell(tr,'No journals match these filters.');td.colSpan=7;td.className='empty-state';tbody.appendChild(tr);return;}for(const j of rows){const tr=document.createElement('tr');const td=makeCell(tr,'');const name=document.createElement('strong');text(name,j.journal);td.appendChild(name);if(j.link){const line=document.createElement('span');line.className='subline';line.appendChild(makeLink('Scopus source ↗',j.link));td.appendChild(line);}makeCell(tr,j.count);makeCell(tr,j.score===null?'N/A':j.score.toFixed(2));makeCell(tr,j.year||'—');makeCell(tr,j.pct===null?'N/A':j.pct+'th');makeCell(tr,badge(j.quartile));makeCell(tr,j.category||'—');tbody.appendChild(tr);}}
  function renderAll(){renderContext();renderDirectory();renderOverview();updatePaperJournalOptions();renderPapers();renderAuthors();renderJournals();}
  function switchTab(tab){state.tab=tab;document.querySelectorAll('.tab').forEach(el=>{const active=el.dataset.tab===tab;el.classList.toggle('active',active);el.setAttribute('aria-selected',String(active));});document.querySelectorAll('.view').forEach(el=>el.hidden=el.id!=='view-'+tab);}
  function exportSummary(){const data=state.data,year=state.year, rows=C.rankAuthors(data,year).map(x=>[x.author.name,x.author.profile,x.total,...C.QUARTILES.map(q=>x.quartiles[q]),x.author.status]);downloadCsv(`Scopus_Author_Summary_${year||'all_years'}.csv`,['Author Name','Scopus Profile URL','Annual Papers','Q1','Q2','Q3','Q4','N/A','Retrieval Status'],rows);}
  function exportPapers(){const year=state.year;downloadCsv(`Scopus_Papers_${year||'all'}_${state.selectedAuthor?'author':'institute'}.csv`,['Author Name','Paper Title','Publication Date','Month','Journal','CiteScore','CiteScore Year','Quartile','Highest Percentile','DOI','Scopus URL'],filteredPapers().map(p=>[p.authorName,p.title,p.date,p.month,p.journal,p.citeScore??'',p.citeScoreYear??'',p.quartile,p.percentile??'',p.doi,p.scopusUrl]));}
  function exportJournals(){downloadCsv('Scopus_Journal_Metrics.csv',['Journal','Papers','CiteScore','Metric Year','Highest Percentile','Quartile','Best Category','Source URL'],filteredJournals().map(x=>[x.journal,x.count,x.score??'',x.year??'',x.pct??'',x.quartile,x.category,x.link]));}
  function setup(){for(const id of ['xlsxInput','welcomeInput'])$(id).addEventListener('change',e=>parseFiles(e.target.files));
    for(const id of ['demoBtn','welcomeDemo'])$(id).addEventListener('click',()=>loadReports(demoReports(),true));
    $('yearSelect').addEventListener('change',e=>{state.year=e.target.value?Number(e.target.value):null;state.page=1;renderAll();});
    $('authorSearch').addEventListener('input',()=>{renderDirectory();const q=C.norm($('authorSearch').value);const exact=state.data.authors.filter(a=>C.norm(a.name)===q);if(exact.length===1&&exact[0].key!==state.selectedAuthor)selectAuthor(exact[0].key);});
    $('authorSearch').addEventListener('keydown',e=>{if(e.key!=='Enter')return;const q=C.norm($('authorSearch').value);const matching=state.data.authors.filter(a=>C.norm(a.name).includes(q)||a.id.includes(q));if(q&&matching.length){selectAuthor(matching[0].key);e.preventDefault();}});
    $('resetFilters').addEventListener('click',()=>{state.selectedAuthor=null;state.year=state.data.years[0]||null;state.page=1;$('yearSelect').value=String(state.year||'');$('authorSearch').value='';$('paperSearch').value='';$('paperQuartile').value='';$('paperJournal').value='';$('paperSort').value='newest';$('rankingSearch').value='';$('journalSearch').value='';$('journalQuartile').value='';renderAll();});
    document.querySelectorAll('.tab').forEach(x=>x.addEventListener('click',()=>switchTab(x.dataset.tab)));
    for(const id of ['paperSearch','paperQuartile','paperJournal','paperSort'])$(id).addEventListener(id==='paperSearch'?'input':'change',()=>{state.page=1;renderPapers();});
    $('pagePrev').addEventListener('click',()=>{state.page--;renderPapers();});$('pageNext').addEventListener('click',()=>{state.page++;renderPapers();});
    $('rankingSearch').addEventListener('input',renderAuthors);$('journalSearch').addEventListener('input',renderJournals);$('journalQuartile').addEventListener('change',renderJournals);
    $('exportSummaryBtn').addEventListener('click',exportSummary);$('exportAuthors').addEventListener('click',exportSummary);$('exportPapers').addEventListener('click',exportPapers);$('exportJournals').addEventListener('click',exportJournals);
    window.addEventListener('dragover',e=>e.preventDefault());window.addEventListener('drop',e=>{e.preventDefault();if(e.dataTransfer?.files?.length)parseFiles(e.dataTransfer.files);});
  }
  setup();
  loadOptionalPublishedReport();
  window.__scopusDashboardTest={loadReports, demoReports, state,core:C}; // browser-only validation helper
})();
