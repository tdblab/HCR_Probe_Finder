"""Build HCR_probe_finder.html from every designer's workbook in a folder.

Usage:  python build_viewer.py [data_folder] [output.html]
        data_folder defaults to "data" next to this script; output defaults to "index.html" next to this script,
        which is the file GitHub Pages serves.

Every .xlsx in the data folder is read as one designer's file, except files whose names start with "_" and Excel
lock files ("~$..."). gene_names.xlsx next to this script supplies short and long gene names. Save each workbook
in Excel before building so that formula results are stored in the file."""
import sys, os, glob, json, re, datetime
from collections import Counter, defaultdict
from openpyxl import load_workbook

TEMPLATE = r'''<!doctype html>
<html lang="en">
<head>
<meta charset="utf-8">
<meta name="viewport" content="width=device-width, initial-scale=1">
<title>HCR probe finder</title>
<style>
:root{
  --paper:#F5F6F4; --sheet:#FFFFFF; --ink:#18202B; --muted:#5B6573; --rule:#D8DCE1; --soft:#EEF0EE;
  --b1:#D9791F; --b2:#B0224F; --b3:#2B8F55; --b4:#A88A12; --b5:#C9481F;
  --warn:#8A4B00; --warnbg:#FFF4E2; --ok:#2B8F55;
  --sans:"Inter","Segoe UI",system-ui,-apple-system,Roboto,Helvetica,Arial,sans-serif;
  --mono:ui-monospace,"Cascadia Mono","SF Mono",Menlo,Consolas,"Courier New",monospace;
}
*{box-sizing:border-box}
html,body{margin:0;height:100%}
body{font:15px/1.5 var(--sans);color:var(--ink);background:var(--paper);display:flex;flex-direction:column}
button,input,select{font:inherit;color:inherit}
:focus-visible{outline:2px solid var(--ink);outline-offset:2px}
header{padding:20px 28px 14px;border-bottom:1px solid var(--rule);background:var(--sheet)}
.top{display:flex;align-items:baseline;justify-content:space-between;gap:16px;flex-wrap:wrap}
h1{font-size:26px;line-height:1.1;margin:0;font-weight:700;letter-spacing:-.01em}
.stamp{color:var(--muted);font-size:13px}
.controls{display:flex;gap:10px;margin-top:14px;flex-wrap:wrap;align-items:center}
#q,#tq{flex:1 1 320px;min-width:220px;padding:10px 14px;border:1px solid var(--rule);border-radius:8px;background:var(--paper);font-size:16px}
#q:focus,#tq:focus{background:#fff}
select{padding:9px 10px;border:1px solid var(--rule);border-radius:8px;background:#fff;max-width:260px}
.ampf{display:flex;gap:6px;flex-wrap:wrap}
.ampf button{border:1px solid var(--rule);background:#fff;border-radius:999px;padding:6px 11px;cursor:pointer;font-size:13px;display:flex;gap:6px;align-items:center}
.ampf button[aria-pressed="true"]{border-color:var(--ink);background:var(--ink);color:#fff}
.dot{width:10px;height:10px;border-radius:50%;display:inline-block;flex:none}
main{flex:1;display:grid;grid-template-columns:minmax(260px,340px) 1fr;min-height:0}
#list{border-right:1px solid var(--rule);overflow:auto;background:var(--sheet)}
#count,.count{padding:10px 18px;color:var(--muted);font-size:13px;border-bottom:1px solid var(--rule);position:sticky;top:0;background:var(--sheet)}
.g{display:block;width:100%;text-align:left;border:0;border-bottom:1px solid var(--soft);background:none;padding:10px 18px;cursor:pointer}
.g:hover{background:var(--paper)}
.g[aria-current="true"]{background:var(--paper);box-shadow:inset 3px 0 0 var(--ink)}
.g .n{font-weight:600}
.g .ln{font-weight:400;font-style:italic}
.gh .lng{font-weight:400;font-style:italic;font-size:.8em}
.g .m{color:var(--muted);font-size:13px;display:flex;gap:8px;align-items:center;flex-wrap:wrap}
.chips{display:inline-flex;gap:3px}
#detail{overflow:auto;padding:24px 32px 120px}
.empty{color:var(--muted);max-width:56ch;margin-top:40px}
.empty h2{color:var(--ink);font-size:20px;margin:0 0 8px}
.gh h2{font-size:30px;margin:0;line-height:1.15;font-weight:700}
.gh .sub{color:var(--muted);margin-top:4px}
.gh .alias{margin-top:6px;font-size:14px}
.set{margin-top:28px;border-top:3px solid var(--c);padding-top:14px}
.sethead{display:flex;justify-content:space-between;gap:16px;flex-wrap:wrap;align-items:flex-start}
.sethead h3{margin:0;font-size:18px;display:flex;gap:10px;align-items:center;flex-wrap:wrap}
.chan{display:inline-flex;align-items:center;gap:6px;padding:2px 10px;border-radius:999px;background:var(--c);color:#fff;font-size:13px;font-weight:600}
.facts{display:grid;grid-template-columns:repeat(auto-fill,minmax(230px,1fr));gap:4px 20px;margin:10px 0 0;font-size:14px}
.facts div span{color:var(--muted)}
.note{margin-top:10px;padding:8px 12px;background:var(--warnbg);color:var(--warn);border-radius:6px;font-size:14px;max-width:90ch}
.actions{display:flex;gap:8px;align-items:center;flex-wrap:wrap}
.btn{text-decoration:none;display:inline-block;border:1px solid var(--ink);background:var(--ink);color:#fff;border-radius:8px;padding:7px 12px;cursor:pointer;font-size:14px}
.btn.ghost{background:#fff;color:var(--ink);border-color:var(--rule)}
.btn:disabled{opacity:.45;cursor:default}
label.sw{font-size:14px;color:var(--muted);display:flex;gap:6px;align-items:center}
.retarget{font-size:13px;color:var(--warn);margin-top:6px}
table{border-collapse:collapse;width:100%;margin-top:12px;font-size:13px}
th{text-align:left;font-weight:600;color:var(--muted);border-bottom:1px solid var(--rule);padding:6px 8px;white-space:nowrap}
td{border-bottom:1px solid var(--soft);padding:6px 8px;vertical-align:top}
td.num{color:var(--muted);width:2.5em}
.oname{font-size:12px;color:var(--muted)}
.seq{font-family:var(--mono);font-size:12.5px;letter-spacing:.02em;word-break:break-all}
.seq .i{color:var(--c);font-weight:700}
.seq .s{color:#9AA3AD}
.flag{color:var(--warn);font-size:12px}
.gm{font-size:12px;white-space:nowrap}
.gm.ok{color:var(--ok)}
.gm.no{color:var(--warn)}
.legend{display:flex;gap:16px;font-size:13px;color:var(--muted);margin-top:18px;flex-wrap:wrap}
.legend .seq{font-size:13px}
#tray{position:fixed;left:0;right:0;bottom:0;background:var(--ink);color:#fff;padding:10px 28px;display:none;gap:14px;align-items:center;flex-wrap:wrap;z-index:5}
#tray.on{display:flex}
#tray .items{display:flex;gap:6px;flex-wrap:wrap;flex:1}
#tray .it{display:inline-flex;align-items:center;gap:6px;border:1px solid #3A4554;border-radius:999px;padding:3px 6px 3px 10px;font-size:13px}
#tray .it button{border:0;background:none;color:#B9C2CC;cursor:pointer;font-size:15px;line-height:1;padding:0 4px}
#tray input{background:#26303D;border:1px solid #3A4554;border-radius:6px;padding:6px 8px;color:#fff;width:180px}
#tray .btn{background:#fff;color:var(--ink);border-color:#fff}
#tray .btn.ghost{background:transparent;color:#fff;border-color:#3A4554}
#trayWarn{color:#FFC98A;font-size:13px;width:100%}
.back{display:none}
.tabs{display:flex;gap:4px;margin-top:12px;border-bottom:1px solid var(--rule)}
.tabs button{border:0;background:none;padding:8px 14px;cursor:pointer;font-size:15px;color:var(--muted);border-bottom:3px solid transparent;margin-bottom:-1px}
.tabs button[aria-selected="true"]{color:var(--ink);border-bottom-color:var(--ink);font-weight:600}
[hidden]{display:none!important}
.iso{width:auto;min-width:60%}
.isowrap{max-width:100%;overflow-x:auto}
.iso td,.iso th{white-space:nowrap}
.iso td.u{white-space:normal;color:var(--muted);font-size:12px;max-width:420px}
.opts{display:flex;gap:16px;flex-wrap:wrap;align-items:center;margin:14px 0 4px;font-size:14px}
.opts input[type=number]{width:70px;padding:6px 8px;border:1px solid var(--rule);border-radius:6px}
.radio{display:flex;gap:12px;flex-wrap:wrap}
.hint{color:var(--muted);font-size:13px;max-width:90ch}
.amp-table td{vertical-align:top}
.amp-table .seq{font-size:12px}
.copy{border:1px solid var(--rule);background:#fff;border-radius:6px;padding:1px 7px;font-size:12px;cursor:pointer;margin-left:6px}
#ampView,#guideView,#submitView{display:block;overflow:auto;padding:24px 32px 120px}
.dl{display:inline-flex;align-items:center;gap:8px;text-decoration:none;font-size:16px;padding:11px 18px;margin-top:4px}
.mock{width:auto;font-size:13px;margin-top:4px;background:#fff;border:1px solid var(--rule)}
.mock th{background:var(--ink);color:#fff;font-weight:600}
.mock td{border:1px solid var(--rule)}
.mock td.y{background:#FFF3B0}.mock td.gr{background:#E6E8EB;color:#9AA3AD}.mock td.ok{background:#DDF2E4}
.mailbox{margin-top:12px;display:flex;gap:10px;align-items:center;flex-wrap:wrap}
.mailbox .addr{font:600 18px var(--mono);background:var(--sheet);border:1px solid var(--rule);border-radius:8px;padding:8px 12px}
@media (max-width:820px){#ampView,#guideView,#submitView{padding:18px 18px 120px}}
.guide{max-width:1100px;margin:0 auto}
.guide h2{font-size:28px;margin:0 0 4px}
.guide .lead{color:var(--muted);margin:0 0 8px;font-size:16px}
.step{margin-top:34px;padding-top:18px;border-top:1px solid var(--rule)}
.step h3{font-size:20px;margin:0 0 12px;display:flex;align-items:center;gap:10px}
.step h3 .k{display:inline-grid;place-items:center;width:30px;height:30px;border-radius:50%;background:var(--ink);color:#fff;font-size:15px}
.step img{width:100%;height:auto;border:1px solid var(--rule);border-radius:8px;display:block;background:#fff}
.caps{display:grid;grid-template-columns:repeat(auto-fill,minmax(220px,1fr));gap:6px 18px;margin:12px 0 0;padding:0;list-style:none;counter-reset:c}
.caps li{counter-increment:c;display:flex;gap:8px;align-items:baseline;font-size:15px}
.caps li::before{content:counter(c);display:inline-grid;place-items:center;min-width:22px;height:22px;border-radius:50%;background:var(--ink);color:#fff;font-size:12px;font-weight:700;flex:none}
.tips{display:grid;grid-template-columns:repeat(auto-fill,minmax(240px,1fr));gap:12px;margin-top:12px}
.tips div{background:var(--sheet);border:1px solid var(--rule);border-radius:8px;padding:12px 14px;font-size:15px}
.tips b{display:block;margin-bottom:2px}
.diagram{background:#fff;border:1px solid var(--rule);border-radius:8px;padding:10px}
.diagram svg{width:100%;height:auto;display:block}
.prose{max-width:90ch}
@media (max-width:820px){
  main{grid-template-columns:1fr}
  #list{border-right:0}
  body.showing #list,body.showing #tlist{display:none}
  body:not(.showing) #detail,body:not(.showing) #tdetail{display:none}
  .back{display:inline-block;margin-bottom:12px}
  #detail{padding:18px 18px 140px}
  header{padding:16px 18px 12px}
  .gh h2{font-size:24px}
  .tabs{flex-wrap:wrap;gap:0 2px}
  .tabs button{padding:7px 9px;font-size:14px}
  select{max-width:100%}
}
@media (prefers-reduced-motion:no-preference){ .g,.btn{transition:background-color .12s} }
</style>
</head>
<body>
<header>
  <div class="top">
    <h1>HCR probe finder</h1>
    <div class="stamp" id="stamp"></div>
  </div>
  <nav class="tabs" role="tablist">
    <button role="tab" data-tab="lab" aria-selected="true">Lab probe sets</button>
    <button role="tab" data-tab="tx" aria-selected="false">Design from transcriptome</button>
    <button role="tab" data-tab="submit" aria-selected="false">Submit your probes</button>
    <button role="tab" data-tab="amp" aria-selected="false">Amplifiers &amp; hairpins</button>
    <button role="tab" data-tab="guide" aria-selected="false">How to use</button>
  </nav>
  <div class="controls" id="labControls">
    <input id="q" type="search" placeholder="Search a gene by short or long name, protein name, LOC number, designer or location" aria-label="Search" autocomplete="off">
    <select id="org" aria-label="Organism"></select>
    <select id="by" aria-label="Designed by"></select>
    <select id="loc" aria-label="Location" hidden></select>
    <div class="ampf" id="ampf" role="group" aria-label="Amplifier"></div>
  </div>
  <div class="controls" id="txControls" hidden>
    <input id="tq" type="search" placeholder="Search any gene: symbol, protein name, LOC / GeneID, or lab name" aria-label="Search transcriptome genes" autocomplete="off">
    <select id="tsp" aria-label="Species"></select>
  </div>
</header>
<main id="labView">
  <nav id="list" aria-label="Genes"><div id="count"></div><div id="genes"></div></nav>
  <section id="detail" aria-live="polite"></section>
</main>
<main id="txView" hidden>
  <nav id="tlist" aria-label="Transcriptome genes" style="border-right:1px solid var(--rule);overflow:auto;background:var(--sheet)"><div id="tcount" class="count"></div><div id="tgenes"></div></nav>
  <section id="tdetail" aria-live="polite" style="overflow:auto;padding:24px 32px 120px"></section>
</main>
<section id="ampView" hidden></section>
<section id="guideView" hidden></section>
<section id="submitView" hidden></section>
<div id="tray" role="region" aria-label="Order">
  <strong>Order</strong>
  <div class="items" id="trayItems"></div>
  <input id="pool" aria-label="Pool name" placeholder="Pool name">
  <button class="btn" id="dlPool">Download oPool sheet</button>
  <button class="btn ghost" id="dlList">Download oligo list</button>
  <button class="btn ghost" id="dlHair">Download hairpin list</button>
  <button class="btn ghost" id="clear">Clear</button>
  <div id="trayWarn"></div>
</div>
<script>
const D=/*__DATA__*/null;
const AMPS=Object.keys(D.amps), COLOR={B1:'var(--b1)',B2:'var(--b2)',B3:'var(--b3)',B4:'var(--b4)',B5:'var(--b5)'};
const st={q:'',org:'',by:'',loc:'',amps:new Set(),gene:null,over:{},order:[]};
const $=id=>document.getElementById(id);
const esc=s=>String(s??'').replace(/[&<>"']/g,c=>({'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;',"'":'&#39;'}[c]));
const em=t=>`<span class="dot" style="background:${COLOR[t]||'#999'}"></span>`;
const nSets=Object.keys(D.sets).length, nPairs=Object.values(D.sets).reduce((a,s)=>a+s.p.length,0);
$('stamp').textContent=`${D.genes.length} genes, ${nSets} probe sets, ${nPairs.toLocaleString()} pairs. Built ${D.built} from ${D.source}`;
const orgs=[...new Set(Object.values(D.sets).map(s=>s.org))].sort();
$('org').innerHTML='<option value="">All organisms</option>'+orgs.map(o=>`<option>${esc(o)}</option>`).join('');
$('by').innerHTML='<option value="">Anyone</option>'+D.designers.map(o=>`<option>${esc(o)}</option>`).join('');
const locs=[...new Set(Object.values(D.sets).map(s=>s.tube).filter(Boolean))].sort();
if(locs.length){$('loc').innerHTML='<option value="">Any location</option>'+locs.map(o=>`<option>${esc(o)}</option>`).join('');$('loc').hidden=false;}
$('ampf').innerHTML=AMPS.map(a=>`<button type="button" data-a="${a}" aria-pressed="false">${em(a)}${a} <span style="opacity:.7">${esc(D.amps[a].fl)}</span></button>`).join('');
// search index per gene
const hay={};
for(const g of D.genes){
  const ss=g.sets.map(id=>D.sets[id]);
  hay[g.k]=[g.label,g.long,g.aka.join(' '),g.names.join(' '),g.prot,g.loc,g.org,...ss.map(s=>[s.full,s.loc,s.gloc,s.names,s.by,s.sid,s.tube,s.sub?s.sub.lab+' '+s.sub.inst:''].join(' '))].join(' ').toLowerCase();
}
function visibleGenes(){
  const toks=st.q.toLowerCase().split(/\s+/).filter(Boolean);
  return D.genes.filter(g=>{
    const ss=g.sets.map(id=>D.sets[id]);
    if(st.org&&g.org!==st.org)return false;
    if(st.by&&!ss.some(s=>s.by===st.by))return false;
    if(st.loc&&!ss.some(s=>s.tube===st.loc))return false;
    if(st.amps.size&&!ss.some(s=>st.amps.has(s.amp)))return false;
    return toks.every(t=>hay[g.k].includes(t));
  });
}
function renderList(){
  const gs=visibleGenes();
  $('count').textContent=gs.length===D.genes.length?`${gs.length} genes`:`${gs.length} of ${D.genes.length} genes`;
  $('genes').innerHTML=gs.length?gs.map(g=>{
    const ss=g.sets.map(id=>D.sets[id]);
    const amps=[...new Set(ss.map(s=>s.amp))].sort();
    const by=[...new Set(ss.map(s=>s.by))].join(', ');
    const also=g.names.filter(n=>n!==g.label&&n.toLowerCase()!==(g.long||'').toLowerCase());
    return `<button class="g" data-k="${esc(g.k)}" aria-current="${g.k===st.gene}"><div class="n">${esc(g.label)}${g.long?` <span class="ln">${esc(g.long)}</span>`:''}${also.length?` <span style="font-weight:400;color:var(--muted)">also ${esc(also.join(', '))}</span>`:''}</div>
      <div class="m"><span class="chips">${amps.map(em).join('')}</span><span>${esc(g.org)}</span><span>${esc(by)}</span></div></button>`;
  }).join(''):`<p class="empty" style="padding:0 18px">No genes match. Try fewer words, or clear the filters.</p>`;
}
function oligo(seq,amp,part){
  // initiator half (18) + spacer (2) + binding arm
  if(part==='P1')return `<span class="seq" style="--c:${COLOR[amp]}"><span class="i">${seq.slice(0,18)}</span><span class="s">${seq.slice(18,20)}</span>${seq.slice(20)}</span>`;
  const n=seq.length;return `<span class="seq" style="--c:${COLOR[amp]}">${seq.slice(0,n-20)}<span class="s">${seq.slice(n-20,n-18)}</span><span class="i">${seq.slice(n-18)}</span></span>`;
}
function build(s,amp){
  // same binding arms; keep recorded sequences for the original amplifier, rebuild for any other
  const A=D.amps[amp];
  return s.p.map(r=>{
    const [n,n1,a1,n2,a2,q1,q2]=r;
    if(amp===s.amp)return {n,n1,n2,q1,q2};
    const re=new RegExp('_'+s.amp+'_');
    return {n,n1:n1&&n1.replace(re,'_'+amp+'_'),n2:n2&&n2.replace(re,'_'+amp+'_'),
            q1:a1?A.p1+A.s1+a1:'',q2:a2?a2+A.s2+A.p2:''};
  });
}
function subFacts(s){const u=s.sub,f=(k,v)=>v?`<div><span>${k}</span> ${esc(v)}</div>`:'';
  return `<div class="facts">${f('Designed by',s.by)}${f('Lab',[u.lab,u.inst].filter(Boolean).join(', '))}${f('Set',s.sid)}${f('Submitted',u.date)}
    ${f('Gene ID / LOC',s.loc)}${f('Did it work?',s.status)}${f('Tested in',u.tissue)}${f('Binds',u.region)}${f('Transcript',u.tx)}
    ${f('Designed with',u.method)}${f('Ordered',[u.odate,u.vendor].filter(Boolean).join(', '))}${f('Location',s.tube)}</div>`;}
function renderDetail(){
  const g=D.genes.find(x=>x.k===st.gene);
  if(!g){$('detail').innerHTML=`<div class="empty"><h2>Find a probe set</h2><p><a href="#guide">First time here? See How to use.</a></p><p>Search for a gene by its name, the protein name, or its LOC number. Each gene lists every probe set made for it, who designed it, which amplifier it carries, and the full oligo sequences.</p><p>Add sets to an order to download an oPool sheet. A set can be ordered on a different amplifier: the binding arms stay the same and only the initiator changes.</p></div>`;return;}
  const ss=g.sets.map(id=>D.sets[id]).sort((a,b)=>a.amp.localeCompare(b.amp)||a.by.localeCompare(b.by));
  let h=`<button class="btn ghost back" id="back">Back to genes</button>
    <div class="gh"><h2>${esc(g.label)}${g.long?` <span class="lng">${esc(g.long)}</span>`:''}</h2>${g.aka.length?`<div class="sub">Also known as ${esc(g.aka.join(', '))}</div>`:''}<div class="sub">${ss.every(x=>x.sub)?`${g.loc?esc(g.loc)+'. ':''}${esc(g.org)}`:`NCBI: ${esc(g.prot||'no annotated protein')}${g.loc?` <span>(${esc(g.loc)})</span>`:''}. ${esc(g.org)}`}</div>
    ${g.names.length>1?`<div class="alias">Sets for this locus are named ${g.names.map(n=>`<b>${esc(n)}</b>`).join(', ')}.</div>`:''}</div>`;
  for(const s of ss){
    const amp=st.over[s.id]||s.amp, P=build(s,amp), inOrder=st.order.some(o=>o.id===s.id);
    const gv=s.gv===null?(s.checked?'no exact match in genome_cds.fa':'not checked yet'):`${s.gv} of ${s.p.length}`;
    h+=`<article class="set" style="--c:${COLOR[amp]}">
      <div class="sethead"><h3><span class="chan">${amp} ${esc(D.amps[amp].fl)}</span> ${esc(s.gene)} <span style="font-weight:400;color:var(--muted)">${s.p.length} pairs</span></h3>
        <div class="actions"><label class="sw">Order on <select data-over="${s.id}">${AMPS.map(a=>`<option value="${a}" ${a===amp?'selected':''}>${a} (${esc(D.amps[a].fl)})${a===s.amp?', as designed':''}</option>`).join('')}</select></label>
        <button class="btn ghost" data-copy="${s.id}">Copy sequences</button>
        <button class="btn" data-add="${s.id}" ${inOrder?'disabled':''}>${inOrder?'In order':'Add to order'}</button></div></div>
      ${amp!==s.amp?`<div class="retarget">Showing the ${amp} version: same binding arms as the designed ${s.amp} set, with the ${amp} initiator. This version has not been ordered before.</div>`:''}
      ${s.sub?subFacts(s):`<div class="facts"><div><span>Designed by</span> ${esc(s.by)}</div><div><span>Set</span> ${esc(s.sid)}</div>
        <div><span>Genome match</span> ${esc(gv)}</div><div><span>LOC</span> ${esc(s.loc||s.gloc||'none recorded')}</div>
        ${s.status?`<div><span>Status</span> ${esc(s.status)}</div>`:''}${s.tube?`<div><span>Location</span> ${esc(s.tube)}</div>`:''}</div>`}
      ${s.notes?`<div class="facts" style="grid-template-columns:1fr"><div><span>Notes</span> ${esc(s.notes)}</div></div>`:''}
      ${s.check?`<div class="note">${esc(s.check)}</div>`:''}
      <table><thead><tr><th>#</th><th>P1</th><th>P2</th>${s.sub?'':'<th>Genome</th>'}</tr></thead><tbody>
      ${P.map((r,i)=>{const raw=s.p[i];const gm=raw[7];return `<tr><td class="num">${esc(r.n)}</td>
        <td>${r.q1?`<div class="oname">${esc(r.n1)}</div>${oligo(r.q1,amp,'P1')}`:'<span class="flag">missing</span>'}</td>
        <td>${r.q2?`<div class="oname">${esc(r.n2)}</div>${oligo(r.q2,amp,'P2')}`:'<span class="flag">missing</span>'}${raw[8]?`<div class="flag">${esc(raw[8])}</div>`:''}</td>
        ${s.sub?'':`<td class="gm ${gm==='Both halves'?'ok':(gm?'no':'')}">${esc(gm==='Both halves'?'both halves':gm==='None'?'no match':gm)}</td>`}</tr>`}).join('')}
      </tbody></table></article>`;
  }
  h+=`<div class="legend"><span><span class="seq" style="--c:var(--ink)"><span class="i">initiator half</span></span> in the amplifier color</span><span><span class="seq"><span class="s">spacer</span></span></span><span><span class="seq">binding arm</span> in black</span></div>`;
  $('detail').innerHTML=h;
}
const oGene=o=>o.kind==='tx'?o.gene:D.sets[o.id].gene;
const oFl=o=>o.fl||D.amps[o.amp].fl;
function renderTray(){
  const t=$('tray'); t.classList.toggle('on',st.order.length>0);
  $('trayItems').innerHTML=st.order.map((o,i)=>`<span class="it">${em(o.amp)}${esc(oGene(o))} ${o.amp} ${esc(oFl(o))}${o.kind==='tx'?' (designed)':''}<button aria-label="Remove ${esc(oGene(o))}" data-rm="${i}">×</button></span>`).join('');
  const byAmp={},byFl={};
  for(const o of st.order){(byAmp[o.amp]=byAmp[o.amp]||new Set()).add(oGene(o));(byFl[oFl(o)]=byFl[oFl(o)]||new Set()).add(o.amp)}
  const w=Object.entries(byAmp).filter(([a,g])=>g.size>1).map(([a,g])=>`${[...g].join(' and ')} are both on ${a}, so they would show in the same channel`)
    .concat(Object.entries(byFl).filter(([f,a])=>a.size>1).map(([f,a])=>`${[...a].join(' and ')} would both be read out with ${f}`));
  $('trayWarn').textContent=w.join('. ');
  if(!$('pool').value||$('pool').dataset.auto==='1'){$('pool').value=[...new Set(st.order.map(oGene))].join('_').replace(/[^A-Za-z0-9_\-]+/g,'-');$('pool').dataset.auto='1'}
}
function orderRows(){const rows=[];for(const o of st.order){
  if(o.kind==='tx'){for(const r of o.rows)rows.push([r[0],r[1],o.src,o.gene,o.amp]);continue}
  const s=D.sets[o.id];for(const r of build(s,o.amp)){if(r.q1)rows.push([r.n1,r.q1,s.sid,s.gene,o.amp]);if(r.q2)rows.push([r.n2,r.q2,s.sid,s.gene,o.amp]);}}return rows;}
function download(name,text){const a=document.createElement('a');a.href=URL.createObjectURL(new Blob([text],{type:'text/csv'}));a.download=name;a.click();setTimeout(()=>URL.revokeObjectURL(a.href),1000);}
const csv=v=>/[",\n]/.test(v)?`"${v.replace(/"/g,'""')}"`:v;
function render(){renderList();renderDetail();renderTray();document.body.classList.toggle('showing',!!st.gene);}
// events
$('q').addEventListener('input',e=>{st.q=e.target.value;renderList();});
$('org').addEventListener('change',e=>{st.org=e.target.value;renderList();});
$('by').addEventListener('change',e=>{st.by=e.target.value;renderList();});
$('loc').addEventListener('change',e=>{st.loc=e.target.value;renderList();});
$('ampf').addEventListener('click',e=>{const b=e.target.closest('button');if(!b)return;const a=b.dataset.a;st.amps.has(a)?st.amps.delete(a):st.amps.add(a);b.setAttribute('aria-pressed',st.amps.has(a));renderList();});
$('genes').addEventListener('click',e=>{const b=e.target.closest('.g');if(!b)return;location.hash=encodeURIComponent(b.dataset.k);});
$('detail').addEventListener('change',e=>{const id=e.target.dataset.over;if(id){st.over[id]=e.target.value;renderDetail();}});
$('detail').addEventListener('click',e=>{
  const t=e.target;
  if(t.id==='back'){history.pushState('',document.title,location.pathname);st.gene=null;render();return;}
  if(t.dataset.add){const id=t.dataset.add;st.order.push({kind:'lab',id,amp:st.over[id]||D.sets[id].amp});render();}
  if(t.dataset.copy){const s=D.sets[t.dataset.copy],amp=st.over[s.id]||s.amp;const txt=build(s,amp).flatMap(r=>[[r.n1,r.q1],[r.n2,r.q2]]).filter(x=>x[1]).map(x=>x.join('\t')).join('\n');
    navigator.clipboard.writeText(txt).then(()=>{t.textContent='Copied';setTimeout(()=>t.textContent='Copy sequences',1500)},()=>{t.textContent='Copy blocked by browser'});}
});
$('trayItems').addEventListener('click',e=>{const i=e.target.dataset.rm;if(i!==undefined){st.order.splice(+i,1);render();}});
$('pool').addEventListener('input',()=>{$('pool').dataset.auto='0'});
$('clear').addEventListener('click',()=>{st.order=[];$('pool').value='';$('pool').dataset.auto='1';render();});
$('dlPool').addEventListener('click',()=>{const p=$('pool').value||'pool';download(p+'_opool.csv','Pool name,Sequence\n'+orderRows().map(r=>[p,r[1]].map(csv).join(',')).join('\n'));});
$('dlHair').addEventListener('click',()=>{const seen=new Map();for(const o of st.order)seen.set(o.amp+'|'+oFl(o),[o.amp,oFl(o)]);
  const rows=[...seen.values()].flatMap(([a,f])=>{const h=D.hairpins[a]||{};return [[a+'-H1-'+f,h.H1||'',a,'H1',f,"3' dye"],[a+'-H2-'+f,h.H2||'',a,'H2',f,"5' dye"]]});
  download(($('pool').value||'pool')+'_hairpins.csv','Name,Sequence,Amplifier,Hairpin,Fluorophore,Dye position\n'+rows.map(r=>r.map(csv).join(',')).join('\n'));});
$('dlList').addEventListener('click',()=>{const p=$('pool').value||'pool';download(p+'_oligos.csv','Name,Sequence,Set_ID,Gene,Amplifier\n'+orderRows().map(r=>r.map(csv).join(',')).join('\n'));});
function showTab(tab){for(const b of document.querySelectorAll('.tabs button'))b.setAttribute('aria-selected',b.dataset.tab===tab);
  $('labView').hidden=tab!=='lab';$('labControls').hidden=tab!=='lab';$('txView').hidden=tab!=='tx';$('txControls').hidden=tab!=='tx';$('ampView').hidden=tab!=='amp';$('guideView').hidden=tab!=='guide';$('submitView').hidden=tab!=='submit';
  document.body.classList.toggle('showing',tab==='lab'?!!st.gene:tab==='tx'?!!TX.gene:false);}
function fromHash(){const h=decodeURIComponent(location.hash.slice(1));
  if(h==='amps'){showTab('amp');renderAmps();return}
  if(h==='guide'){showTab('guide');renderGuide();return}
  if(h==='submit'){showTab('submit');renderSubmit();return}
  if(h==='tx'||h.startsWith('tx/')){const [,sp,sym]=h.split('/');showTab('tx');txOpen(sp||TX.sp,sym||null);return}
  st.gene=D.genes.some(g=>g.k===h)?h:null;showTab('lab');render();if(st.gene)$('detail').scrollTop=0;}
document.querySelector('.tabs').addEventListener('click',e=>{const b=e.target.closest('button');if(!b)return;
  location.hash=b.dataset.tab==='guide'?'guide':b.dataset.tab==='submit'?'submit':b.dataset.tab==='amp'?'amps':b.dataset.tab==='tx'?('tx/'+TX.sp):'';if(b.dataset.tab==='lab'&&!location.hash){st.gene=null;fromHash()}});
// ---------- transcriptome-wide designs (loaded on demand from tx/<species>/) ----------
const TX={sp:(D.tx[0]||{}).id||'bany',idx:{},shards:{},gene:null,sel:{},mode:{},n:{},amp:{},fl:{},hay:{},parts:{}};
const PART={1:"5′UTR",2:'CDS',4:"3′UTR"};
const partTxt=m=>[1,2,4].filter(b=>m&b).map(b=>PART[b]).join(' + ');
const span=(a,e)=>e>a?`${(a+1).toLocaleString()}–${e.toLocaleString()}`:'none';
const DESIGN_AMPS=['B1','B2','B3'], FLUORS=['AF488','AF546','AF647'];
const loaded={};
function loadScript(src){if(loaded[src])return loaded[src];return loaded[src]=new Promise((ok,no)=>{const s=document.createElement('script');s.src=src;s.onload=ok;s.onerror=()=>{delete loaded[src];no(new Error(src))};document.head.appendChild(s);});}
window.HCRTX_INDEX=(sp,obj)=>{TX.idx[sp]=obj;TX.hay[sp]=obj.g.map(g=>[g[0],g[1],g[2],g[7],g[8],g[10]||''].join(' ').toLowerCase());};
window.HCRTX_SHARD=(sp,n,data)=>{(TX.shards[sp]=TX.shards[sp]||{})[n]=data;};
const spName=sp=>(D.tx.find(x=>x.id===sp)||{}).name||sp;
$('tsp').innerHTML=D.tx.map(x=>`<option value="${x.id}">${esc(x.name)}</option>`).join('');
const rcomp=s=>s.split('').reverse().map(c=>({A:'T',C:'G',G:'C',T:'A'}[c])).join('');
function unpack(b){const bin=atob(b);let bits='';for(let i=0;i<bin.length;i++)bits+=bin.charCodeAt(i).toString(2).padStart(8,'0');let s='';for(let i=0;i<52;i++)s+='ACGT'[parseInt(bits.substr(i*2,2),2)];return s;}
async function txIndex(sp){if(!TX.idx[sp])await loadScript(`tx/${sp}/index.js`);return TX.idx[sp];}
async function txOpen(sp,sym){
  TX.sp=sp;$('tsp').value=sp;
  try{await txIndex(sp);}catch(e){$('tdetail').innerHTML=`<div class="empty"><h2>Designs not found</h2><p>The design files for ${esc(spName(sp))} were not found next to this page. Keep the tx folder in the same folder as index.html.</p></div>`;return}
  TX.gene=sym;txList();await txDetail();document.body.classList.toggle('showing',!!sym);}
function txList(){
  const I=TX.idx[TX.sp];if(!I)return;const toks=$('tq').value.toLowerCase().split(/\s+/).filter(Boolean);const H=TX.hay[TX.sp];
  let hits=[];for(let i=0;i<I.g.length;i++){if(toks.every(t=>H[i].includes(t)))hits.push(i);}
  const q=$('tq').value.toLowerCase().trim();
  const rank=g=>g[0].toLowerCase()===q?0:(g[7]||'').toLowerCase().split(', ').includes(q)?1:(g[2]||'').toLowerCase()===q?2:(g[11]||'').toLowerCase().split(', ').includes(q)?3:4;
  hits.sort((a,b)=>rank(I.g[a])-rank(I.g[b])||I.g[a][0].localeCompare(I.g[b][0]));
  $('tcount').textContent=toks.length?`${hits.length.toLocaleString()} of ${I.g.length.toLocaleString()} genes${hits.length>300?', showing 300':''}`:`${I.g.length.toLocaleString()} genes. Type to search`;
  const show=toks.length?hits.slice(0,300):[];
  $('tgenes').innerHTML=show.map(i=>{const g=I.g[i];return `<button class="g" data-sym="${esc(g[0])}" aria-current="${g[0]===TX.gene}"><div class="n">${esc(g[0])}${g[7]&&g[7]!==g[0]?` <span class="ln">${esc(g[7])}</span>`:''}</div>
    <div class="m"><span>${esc(g[2])}</span></div><div class="m"><span>${g[12]!==undefined?`${g[6]} CDS + ${g[12]} UTR pairs`:`${g[6]} pairs`}</span><span>${g[5]} isoform${g[5]>1?'s':''}</span>${g[9]?`<span>lab has ${g[9]} set${g[9]>1?'s':''}</span>`:''}</div></button>`}).join('');
}
async function txGene(sym){const I=TX.idx[TX.sp];const g=I.g.find(x=>x[0]===sym);if(!g)return null;
  if(!(TX.shards[TX.sp]||{})[g[3]])await loadScript(`tx/${TX.sp}/s${String(g[3]).padStart(4,'0')}.js`);
  return {row:g,d:TX.shards[TX.sp][g[3]][g[4]]};}
function txPairs(key,d){
  const n=d.iso.length,all=(1n<<BigInt(n))-1n;
  const sel=TX.sel[key]||all, mode=TX.mode[key]||'every';
  const parts=TX.parts[key]||2;
  const f=d.p.filter(p=>{const m=BigInt('0x'+p[2]);return (m&sel)===sel&&(mode!=='only'||(m&~sel&all)===0n)&&((p[5]??2)&~parts)===0});
  f.sort((a,b)=>b[1]-a[1]);return f;}
async function txDetail(){
  const box=$('tdetail');const I=TX.idx[TX.sp];
  if(!TX.gene){box.innerHTML=`<div class="empty"><h2>Design probes for any gene</h2><p>Every protein-coding gene in ${esc(spName(TX.sp))} (${I?I.meta.genes.toLocaleString():''} genes, assembly ${esc(I?I.meta.assembly:'')}) has pre-computed probe pairs. Search for a gene, choose the isoforms to target, ${I&&I.meta.utr?'where on the transcript the probes may sit (CDS by default, 5′UTR and 3′UTR on request), ':''}how many pairs, the amplifier and the fluorophore, then add the set to an order.</p><p class="hint">${esc(I?I.meta.rules:'')} These designs are computational and have not been tested; sets made by lab members are on the Lab probe sets tab.</p></div>`;return}
  const r=await txGene(TX.gene);if(!r){box.innerHTML='<div class="empty"><h2>Gene not found</h2></div>';return}
  const {row,d}=r,key=TX.sp+'|'+d.s,n=d.iso.length,all=(1n<<BigInt(n))-1n,sel=TX.sel[key]||all,mode=TX.mode[key]||'every';
  const P=txPairs(key,d),N=Math.min(TX.n[key]||20,P.length),amp=TX.amp[key]||'B1',fl=TX.fl[key]||D.amps[amp].fl;
  const pick=P.slice(0,N).sort((a,b)=>a[3]-b[3]||a[4]-b[4]);
  const A=D.amps[amp];const label=(row[7]||'').split(', ')[0]||d.s;const sym=label.replace(/[^A-Za-z0-9.\-()]+/g,'-');
  const rows=pick.map((p,i)=>{const t=unpack(p[0]);const nn=String(i+1).padStart(2,'0');return {nn,t,q1:A.p1+A.s1+rcomp(t.slice(0,25)),q2:rcomp(t.slice(27))+A.s2+A.p2,n1:`${sym}_${amp}_P1_${nn}`,n2:`${sym}_${amp}_P2_${nn}`,p}});
  TX.cur={key,gene:label,amp,fl,rows};
  const labKey=D.genes.find(g=>g.k===(TX.sp==='bany'?d.s:'GeneID:'+d.id));
  const isoLab=i=>I.meta.utr?d.iso[i][5]:d.iso[i][1]?'isoform '+d.iso[i][1]:d.iso[i][0];
  const covTxt=p=>{const m=BigInt('0x'+p[2]);if(m===all)return n>1?'all isoforms':'';const l=[];for(let i=0;i<n;i++)if(m>>BigInt(i)&1n)l.push(I.meta.utr?d.iso[i][5]:d.iso[i][1]||d.iso[i][0]);return l.join(', ')};
  let h=`<button class="btn ghost back" id="tback">Back to genes</button>
   <div class="gh"><h2>${esc(d.s)}${row[7]&&row[7]!==d.s?` <span class="lng">${esc(row[7])}</span>`:''}</h2>
   ${row[8]?`<div class="sub">${esc(row[8])}</div>`:''}<div class="sub">NCBI: ${esc(d.n)} (${esc(d.s.startsWith('LOC')?d.s:'GeneID:'+d.id)}). ${esc(spName(TX.sp))}</div>
   ${row[11]?`<div class="sub">Also known as ${esc(row[11])}</div>`:''}
   ${labKey?`<div class="alias">The lab already has probe sets for this gene: <a href="#${encodeURIComponent(labKey.k)}">open on the Lab probe sets tab</a>.</div>`:''}</div>`;
  if(!d.p.length){box.innerHTML=h+`<div class="note">No gene-specific pairs could be designed. Almost all of this gene's sequence is shared with other genes (for example transposon-derived or multi-copy genes), so any probe would also bind them.</div>`;return}
  const utr=!!I.meta.utr, uTxt=v=>`${v[2]!==undefined&&utr?partTxt(v[2])+' ':''}${(v[0]+1).toLocaleString()}–${v[1].toLocaleString()}`;
  if(n>1||utr){h+=`<h3 style="margin-top:22px">Isoforms</h3><div class="isowrap"><table class="iso"><thead><tr>${n>1?'<th>Target</th>':''}<th>Isoform</th>${utr?'<th>Transcript</th>':''}<th>Protein</th><th>${utr?'Length':'CDS length'}</th>${utr?'<th>5′UTR</th><th>CDS</th><th>3′UTR</th>':''}${n>1?'<th>Unique to this isoform (nt)</th>':''}</tr></thead><tbody>`+
    d.iso.map((x,i)=>{const u=(d.u[i]||[]);const known=x[8]!==0;return `<tr>${n>1?`<td><input type="checkbox" data-iso="${i}" ${(sel>>BigInt(i))&1n?'checked':''} aria-label="Target ${esc(isoLab(i))}"></td>`:''}<td>${esc(x[1]||'')}${utr&&x[9]?` <span class="oname">variant ${esc(x[9])}</span>`:''}</td>${utr?`<td>${esc(x[5])}</td>`:''}<td>${esc(x[0])}${x[3].length?` <span class="oname">same ${utr?'transcript':'CDS'} as ${esc(x[3].join(', '))}</span>`:''}${x[4]?' <span class="flag">partial</span>':''}</td><td>${x[2].toLocaleString()} nt</td>
      ${utr?(known?`<td class="u">${span(0,x[6])}</td><td class="u">${span(x[6],x[7])}</td><td class="u">${span(x[7],x[2])}</td>`:`<td class="u" colspan="3"><span class="flag">UTRs unknown: CDS only</span></td>`):''}
      ${n>1?`<td class="u">${u.length?u.slice(0,6).map(uTxt).join(', ')+(u.length>6?` and ${u.length-6} more`:''):'none'}</td>`:''}</tr>`}).join('')+`</tbody></table></div>`}
  if(n>1){h+=`
    <div class="radio opts" role="radiogroup" aria-label="Isoform mode"><label><input type="radio" name="mode" value="every" ${mode!=='only'?'checked':''}> Pairs present in every selected isoform</label>
    <label><input type="radio" name="mode" value="only" ${mode==='only'?'checked':''}> Only pairs absent from the unselected isoforms (to detect a specific variant)</label></div>`}
  const parts=TX.parts[key]||2;
  if(utr)h+=`<div class="radio opts" role="group" aria-label="Probe location"><span>Probe location</span>${[2,1,4].map(b=>`<label><input type="checkbox" data-part="${b}" ${parts&b?'checked':''}> ${PART[b]}</label>`).join('')}<span class="hint">CDS by default; tick a UTR to add pairs there. A pair crossing a boundary needs both parts ticked.</span></div>`;
  h+=`<div class="opts"><label>Pairs <input type="number" id="tn" min="1" max="${P.length}" value="${N}"></label><span class="hint">of ${P.length} available${n>1||utr?' for this '+[n>1?'isoform':'',utr?'location':''].filter(Boolean).join(' and ')+' choice':''}</span>
      <label>Amplifier <select id="tamp">${DESIGN_AMPS.map(a=>`<option ${a===amp?'selected':''}>${a}</option>`).join('')}</select></label>
      <label>Fluorophore <select id="tfl">${FLUORS.map(f=>`<option ${f===fl?'selected':''}>${f}</option>`).join('')}</select></label>
      <button class="btn ghost" id="tcopy" ${N?'':'disabled'}>Copy sequences</button><button class="btn" id="tadd" ${N?'':'disabled'}>Add to order</button></div>`;
  if(!N){box.innerHTML=h+`<div class="note">No pairs fit this choice. Try "present in every selected isoform", select fewer isoforms${utr?', or tick another probe location':''}.</div>`;return}
  if(N<5)h+=`<div class="note">Only ${N} pair${N>1?'s':''}: signal may be weak with this few.</div>`;
  h+=`<article class="set" style="--c:${COLOR[amp]}"><div class="sethead"><h3><span class="chan">${amp} ${esc(fl)}</span> ${esc(label)} <span style="font-weight:400;color:var(--muted)">${N} pairs, designed, not tested</span></h3></div>
   <table><thead><tr><th>#</th><th>P1</th><th>P2</th><th>Target</th></tr></thead><tbody>${rows.map(r=>`<tr><td class="num">${r.nn}</td>
   <td><div class="oname">${esc(r.n1)}</div>${oligo(r.q1,amp,'P1')}</td><td><div class="oname">${esc(r.n2)}</div>${oligo(r.q2,amp,'P2')}</td>
   <td class="gm">${utr?`<b>${partTxt(r.p[5]??2)}</b> `:''}nt ${r.p[4]+1}–${r.p[4]+52}${n>1?` of ${esc(isoLab(r.p[3]))}`:''}<div class="oname">quality ${r.p[1]}${n>1?'; '+esc(covTxt(r.p)):''}</div></td></tr>`).join('')}</tbody></table></article>
   <p class="hint" style="margin-top:14px">${esc(I.meta.rules)} Quality is 100 for two arms at exactly 50% GC with no penalties.</p>`;
  box.innerHTML=h;
}
$('tq').addEventListener('input',txList);
$('tsp').addEventListener('change',e=>{location.hash='tx/'+e.target.value;});
$('tgenes').addEventListener('click',e=>{const b=e.target.closest('.g');if(b)location.hash=`tx/${TX.sp}/${encodeURIComponent(b.dataset.sym)}`;});
$('tdetail').addEventListener('change',e=>{const c=TX.cur;if(!c&&!e.target.dataset.iso&&!e.target.dataset.part)return;const key=TX.sp+'|'+TX.gene;const t=e.target;
  if(t.dataset.iso!==undefined){const i=BigInt(t.dataset.iso);let m=TX.sel[key];if(m===undefined){m=0n;for(const x of document.querySelectorAll('[data-iso]'))if(x.checked||x===t)m|=1n<<BigInt(x.dataset.iso);if(!t.checked)m&=~(1n<<i)}else m=t.checked?m|(1n<<i):m&~(1n<<i);if(m===0n){t.checked=true;return}TX.sel[key]=m;}
  if(t.dataset.part!==undefined){const b=+t.dataset.part;let m=TX.parts[key]||2;m=t.checked?m|b:m&~b;if(!m){t.checked=true;return}TX.parts[key]=m;}
  if(t.name==='mode')TX.mode[key]=t.value;
  if(t.id==='tn')TX.n[key]=Math.max(1,parseInt(t.value)||1);
  if(t.id==='tamp'){TX.amp[key]=t.value;if(!TX.fl[key])TX.fl[key]=D.amps[t.value].fl;}
  if(t.id==='tfl')TX.fl[key]=t.value;
  txDetail();});
$('tdetail').addEventListener('click',e=>{const t=e.target,c=TX.cur;
  if(t.id==='tback'){location.hash='tx/'+TX.sp;return}
  if(t.id==='tadd'&&c){st.order.push({kind:'tx',gene:c.gene,amp:c.amp,fl:c.fl,src:'designed ('+spName(TX.sp)+')',rows:c.rows.flatMap(r=>[[r.n1,r.q1],[r.n2,r.q2]])});renderTray();t.textContent='Added';}
  if(t.id==='tcopy'&&c){navigator.clipboard.writeText(c.rows.flatMap(r=>[[r.n1,r.q1],[r.n2,r.q2]]).map(x=>x.join('\t')).join('\n')).then(()=>{t.textContent='Copied'},()=>{t.textContent='Copy blocked by browser'});}});
// ---------- amplifiers and hairpins ----------
function renderAmps(){
  const H=D.hairpins,keys=Object.keys(H).sort((a,b)=>parseInt(a.slice(1))-parseInt(b.slice(1)));
  const cell=s=>`<span class="seq">${esc(s)}</span><button class="copy" data-copyseq="${esc(s)}">Copy</button>`;
  $('ampView').innerHTML=`<div class="prose"><h2 style="margin:0 0 6px">HCR amplifiers and hairpins</h2>
   <p class="hint">Each amplifier is a pair of dye-labeled hairpins, H1 (dye at the 3′ end) and H2 (dye at the 5′ end), opened by the initiator that the split probe pair assembles. I1 and I2 are the two forms of the full initiator. The fluorophore is set by the hairpins, so any amplifier can carry any dye; the lab default is shown for B1–B5. Probe pairs carry half of I2 each: the P1 half is the first 18 nt of I2, the P2 half the last 18 nt, each joined to the binding arm by a 2-nt spacer.</p></div>
   <table class="amp-table"><thead><tr><th>Amplifier</th><th>Lab default dye</th><th>Dye in source document</th><th>I1</th><th>I2</th><th>H1 (3′ dye)</th><th>H2 (5′ dye)</th></tr></thead><tbody>${keys.map(a=>{const h=H[a];return `<tr>
     <td>${COLOR[a]?em(a):''} <b>${a}</b></td><td>${esc((D.amps[a]||{}).fl||'')}</td><td>${esc(h.dye||'')}</td><td>${cell(h.I1)}</td><td>${cell(h.I2)}</td><td>${cell(h.H1)}</td><td>${cell(h.H2)}</td></tr>`}).join('')}</tbody></table>
   <h3 style="margin-top:26px">Initiator halves and spacers used in probes</h3><table class="amp-table" style="width:auto"><thead><tr><th>Amplifier</th><th>P1: initiator half + spacer + arm</th><th>P2: arm + spacer + initiator half</th></tr></thead><tbody>
   ${Object.keys(D.amps).map(a=>{const x=D.amps[a];return `<tr><td>${em(a)} <b>${a}</b></td><td><span class="seq" style="--c:${COLOR[a]}"><span class="i">${x.p1}</span><span class="s">${x.s1}</span>NNN…</span></td><td><span class="seq" style="--c:${COLOR[a]}">…NNN<span class="s">${x.s2}</span><span class="i">${x.p2}</span></span></td></tr>`}).join('')}</tbody></table>
   <p class="hint">${esc(D.hairpinNote||'')}</p>`;
}
$('ampView').addEventListener('click',e=>{const s=e.target.dataset.copyseq;if(s)navigator.clipboard.writeText(s).then(()=>{e.target.textContent='Copied';setTimeout(()=>e.target.textContent='Copy',1200)});});

// ---------- how to use ----------
const DIAGRAM=`<svg viewBox="0 0 1000 300" role="img" aria-label="How a split probe pair triggers amplification">
<style>.t{font:600 15px system-ui,sans-serif;fill:#18202B}.m{font:13px system-ui,sans-serif;fill:#5B6573}</style>
<text class="t" x="20" y="28">1. Two probes land side by side</text>
<line x1="20" y1="190" x2="300" y2="190" stroke="#18202B" stroke-width="5" stroke-linecap="round"/><text class="m" x="20" y="215">mRNA</text>
<line x1="60" y1="172" x2="155" y2="172" stroke="#18202B" stroke-width="5"/><line x1="165" y1="172" x2="260" y2="172" stroke="#18202B" stroke-width="5"/>
<path d="M60 172 L35 110" stroke="#D9791F" stroke-width="5" fill="none" stroke-linecap="round"/><path d="M260 172 L285 110" stroke="#D9791F" stroke-width="5" fill="none" stroke-linecap="round"/>
<text class="m" x="90" y="160">P1 · 25 nt</text><text class="m" x="190" y="160">P2 · 25 nt</text><text class="m" x="145" y="235">2-nt gap</text>
<text class="m" x="20" y="98" fill="#D9791F">initiator half</text>
<text class="t" x="360" y="28">2. The halves make one initiator</text>
<path d="M380 120 Q470 70 560 120" stroke="#D9791F" stroke-width="6" fill="none" stroke-linecap="round"/>
<text class="m" x="405" y="150">only when both probes bind</text><text class="m" x="415" y="170">the same transcript</text>
<path d="M600 110 l30 0 m-10 -10 l10 10 l-10 10" stroke="#5B6573" stroke-width="3" fill="none"/>
<text class="t" x="660" y="28">3. Hairpins H1 and H2 chain up</text>
<g stroke-width="6" fill="none" stroke-linecap="round">
<path d="M670 200 l40 -50 l40 50" stroke="#18202B"/><path d="M750 200 l40 -50 l40 50" stroke="#5B6573"/><path d="M830 200 l40 -50 l40 50" stroke="#18202B"/><path d="M910 200 l30 -38" stroke="#5B6573"/></g>
<g fill="#D9791F"><circle cx="710" cy="150" r="9"/><circle cx="790" cy="150" r="9"/><circle cx="870" cy="150" r="9"/><circle cx="940" cy="162" r="9"/></g>
<text class="m" x="670" y="235">each hairpin carries a dye</text><text class="m" x="670" y="255">so one transcript becomes a bright dot</text>
<text class="m" x="20" y="285">The initiator decides the amplifier (B1, B2, B3…). The dye on the hairpins decides the color (AF488, AF546, AF647).</text></svg>`;
function renderGuide(){
  const img=k=>D.guide&&D.guide[k]?`<img src="${D.guide[k]}" alt="" loading="lazy">`:'';
  $('guideView').innerHTML=`<div class="guide">
   <h2>How to use</h2><p class="lead">Find the lab's probes, design new ones for any gene, and download a ready-to-order sheet.</p>
   <section class="step"><h3><span class="k">1</span>Find a probe set the lab already has</h3>${img('lab')}
    <ol class="caps"><li>Type a gene: short name, full name or LOC</li><li>Click the gene</li><li>Pick the amplifier to order it on</li><li>Add it to your order</li></ol></section>
   <section class="step"><h3><span class="k">2</span>Design probes for any gene</h3>${img('tx')}
    <ol class="caps"><li>Open Design from transcriptome</li><li>Choose the species</li><li>Search for the gene</li><li>Keep all isoforms ticked</li><li>Choose how many pairs</li><li>Choose the amplifier</li><li>Choose the color</li><li>Add it to your order</li></ol></section>
   <section class="step"><h3><span class="k">3</span>Detect one splice variant</h3>${img('iso')}
    <ol class="caps"><li>Tick only the variant you want</li><li>Choose "Only pairs absent from the unselected isoforms"</li><li>Check how many pairs are left</li></ol></section>
   <section class="step"><h3><span class="k">4</span>Download your order</h3>${img('order')}
    <ol class="caps"><li>oPool sheet, ready for IDT</li><li>All oligo names and sequences</li><li>The hairpins you need</li><li>Warnings if two genes share a channel</li></ol></section>
   <section class="step"><h3><span class="k">5</span>How the probes work</h3><div class="diagram">${DIAGRAM}</div>
    <p class="lead" style="margin-top:12px">In every sequence on this site: <span class="seq" style="--c:var(--b1)"><span class="i">initiator half</span></span> in the amplifier color, <span class="seq"><span class="s">spacer</span></span> in gray, <span class="seq">binding arm</span> in black.</p></section>
   <section class="step"><h3><span class="k">6</span>Good to know</h3><div class="tips">
    <div><b>One amplifier per gene</b>Genes on the same amplifier show up in the same channel.</div>
    <div><b>Enough pairs</b>Fewer than 5 pairs may give a weak signal.</div>
    <div><b>Designed sets are untested</b>Lab sets have been made by lab members; designed sets come straight from the computer.</div>
    <div><b>Hairpin sequences</b>All 20 amplifiers are on the Amplifiers &amp; hairpins tab.</div>
    <div><b>Finding the tubes</b>Each lab set shows where it is stored. Use the location menu to list everything in one place.</div>
    <div><b>CDS or UTRs</b>Designs sit in the CDS by default. For species with full transcripts, tick 5′UTR or 3′UTR to add pairs there.</div></div></section>
   <section class="step"><h3><span class="k">7</span>Share your own probes</h3>
    <ol class="caps"><li>Open <a href="#submit">Submit your probes</a></li><li>Download the form</li><li>Fill it in</li><li>Email it</li></ol></section>
  </div>`;
}
// ---------- submit your probes ----------
// The address is stored encoded and only put together when someone clicks, so bots that harvest addresses
// from web pages do not find it.
const MAILK=[116,118,106,53,115,112,104,116,110,71,62,56,105,104,111,123,121,112,123];
const mailAddr=()=>MAILK.map(c=>String.fromCharCode(c-7)).reverse().join('');
function renderSubmit(){
  const mock=`<div class="isowrap"><table class="mock"><thead><tr><th>Gene #</th><th>Gene short name *</th><th>Gene long name *</th><th>Species *</th><th>Gene ID or LOC *</th><th>Amplifier *</th><th>Did it work? *</th><th>Row check</th></tr></thead><tbody>
    <tr><td>1</td><td>hth</td><td>homothorax</td><td>Bicyclus anynana</td><td>LOC…</td><td>B1</td><td>Works well</td><td class="ok">Ready ✓</td></tr>
    <tr><td>2</td><td>Dll</td><td class="y"></td><td>Bicyclus anynana</td><td class="y"></td><td>B2</td><td class="y"></td><td style="background:#FFE2C2">Missing: long name, Gene ID…</td></tr>
    <tr><td class="gr">3</td><td class="gr"></td><td class="gr"></td><td class="gr"></td><td class="gr"></td><td class="gr"></td><td class="gr"></td><td class="gr"></td></tr></tbody></table></div>`;
  const pmock=`<div class="isowrap"><table class="mock"><thead><tr><th>Gene #</th><th>Gene</th><th>Pair #</th><th>P1 sequence (5′→3′) *</th><th>P2 sequence (5′→3′) *</th><th>Check</th></tr></thead><tbody>
    <tr><td>1</td><td>hth</td><td>1</td><td><span class="seq" style="--c:var(--b1)"><span class="i">GAGGAGGGCAGCAAACGG</span><span class="s">AA</span>NNNN…</span></td><td><span class="seq" style="--c:var(--b1)">…NNNN<span class="s">TA</span><span class="i">GAAGAGTCTTCCTTTACG</span></span></td><td class="ok">OK ✓ (B1)</td></tr>
    <tr><td>1</td><td>hth</td><td>2</td><td><span class="seq">…</span></td><td><span class="seq">…</span></td><td class="ok">OK ✓ (B1)</td></tr></tbody></table></div>`;
  $('submitView').innerHTML=`<div class="guide">
   <h2>Submit your probes</h2><p class="lead">Designed and ordered HCR probes? Share them here so others can find and reuse them.</p>
   <section class="step"><h3><span class="k">1</span>Download the form</h3>
    <a class="btn dl" href="templates/HCR_probe_submission.xlsx" download>Download the submission form (.xlsx)</a>
    <p class="hint" style="margin-top:8px">Opens in Excel, Google Sheets, LibreOffice or Numbers. No macros. Up to 30 genes per file.</p></section>
   <section class="step"><h3><span class="k">2</span>Fill it in</h3>
    <ol class="caps"><li><span><b>About you</b>: name, initials, lab, institution</span></li><li><span><b>Genes</b>: type how many, then one row per gene</span></li><li><span><b>Probes</b>: one row per pair: gene #, P1, P2</span></li><li><span>Every pair should say <b>OK ✓</b></span></li></ol>
    <p class="hint" style="margin-top:14px">Yellow cells are required. Rows past your number of genes stay gray.</p>${mock}
    <p class="hint" style="margin-top:14px">Paste whole oligos, 5′ to 3′, exactly as ordered. The sheet reads the amplifier from the initiator.</p>${pmock}</section>
   <section class="step"><h3><span class="k">3</span>Email it</h3>
    <p class="lead">Save the file and attach it to an email.</p>
    <div class="mailbox" id="mailbox"><button class="btn" id="showMail">Show the email address</button></div></section>
   <section class="step"><h3><span class="k">4</span>It appears on the site</h3>
    <p class="lead">Once checked, your sets are added to <a href="#">Lab probe sets</a> under your name, with your lab and how well they worked.</p></section>
   <section class="step"><h3><span class="k">5</span>Good to know</h3><div class="tips">
    <div><b>It becomes public</b>Gene names, sequences, your name and lab are shown to anyone. Don't send sets you want to keep private.</div>
    <div><b>Did it work?</b>Say so honestly, even if it failed. That is what helps the next person most.</div>
    <div><b>More genes or a fix</b>Send another file any time. For a correction, say which set it replaces.</div>
    <div><b>Any amplifier</b>B1 to B5 are recognized from the sequence.</div></div></section>
  </div>`;
}
$('submitView').addEventListener('click',e=>{const t=e.target;
  if(t.id==='showMail'){const a=mailAddr(),sub=encodeURIComponent('HCR probe submission'),body=encodeURIComponent('Hello,\n\nPlease find my HCR probe submission form attached.\n\n');
    $('mailbox').innerHTML=`<span class="addr">${esc(a)}</span><a class="btn" href="mailto:${a}?subject=${sub}&body=${body}">Write the email</a><button class="btn ghost" id="copyMail">Copy address</button>
      <p class="hint" style="width:100%;margin:4px 0 0">Subject: <b>HCR probe submission</b>. Remember to attach the file.</p>`;}
  if(t.id==='copyMail')navigator.clipboard.writeText(mailAddr()).then(()=>{t.textContent='Copied'},()=>{t.textContent='Copy blocked by browser'});});
window.addEventListener('hashchange',fromHash);
fromHash();
</script>
</body>
</html>
'''

here = os.path.dirname(os.path.abspath(__file__))
folder = sys.argv[1] if len(sys.argv) > 1 else os.path.join(here, 'data')
out = sys.argv[2] if len(sys.argv) > 2 else os.path.join(here, 'index.html')
os.makedirs(os.path.dirname(os.path.abspath(out)), exist_ok=True)
files = sorted(f for f in glob.glob(os.path.join(folder, '*.xlsx'))
               if not os.path.basename(f).startswith(('_', '~$')))
if not files: sys.exit(f'No designer workbooks found in {folder}')

def table(ws):
    it = ws.iter_rows(values_only=True); hdr = list(next(it))
    return [dict(zip(hdr, r)) for r in it if r and r[0] not in (None, '')]
def s(v): return '' if v is None else str(v).strip()
def is_example(sid, gene): return sid.upper().startswith('XX-') or gene.lower().startswith('example')

# short and long gene names
NAMES = {}
gpath = os.path.join(here, 'gene_names.xlsx')
if os.path.exists(gpath):
    for r in table(load_workbook(gpath, data_only=True, read_only=True)['Gene_names']):
        if s(r.get('Short name')) or s(r.get('Long name')):
            NAMES[s(r['Name used in files'])] = (s(r.get('Short name')), s(r.get('Long name')), s(r.get('Also known as')))
NAMES_LC = {k.lower(): v for k, v in NAMES.items()}
def names_for(label): return NAMES.get(label) or NAMES_LC.get(label.lower()) or ('', '', '')

amps = None; S = []; designers = []
for f in files:
    wb = load_workbook(f, data_only=True, read_only=True)
    stem = os.path.splitext(os.path.basename(f))[0]
    if amps is None and 'Amplifiers' in wb.sheetnames:
        amps = {r['Amplifier']: dict(fl=r['Fluorophore (lab default)'], p1=r['P1 5′ initiator half'], s1=r['P1 spacer'],
                                     s2=r['P2 spacer'], p2=r['P2 3′ initiator half']) for r in table(wb['Amplifiers'])}
    pairs = defaultdict(list)
    for r in table(wb['Probe_Pairs']):
        pairs[s(r['Set_ID'])].append(r)
    for r in table(wb['Probe_Sets']):
        sid = s(r['Set_ID']); gene = s(r['Gene (short)'])
        if is_example(sid, gene): continue
        by = s(r.get('Designed by')) or stem
        if by not in designers: designers.append(by)
        P = []
        for p in sorted(pairs.get(sid, []), key=lambda p: int(p['Pair #'] or 0)):
            q1, q2 = s(p['P1 sequence (5′→3′)']).upper(), s(p['P2 sequence (5′→3′)']).upper()
            a1 = s(p['P1 binding arm']).upper() or (q1[20:] if len(q1) > 20 else '')   # arm = after 18-nt initiator half + 2-nt spacer
            a2 = s(p['P2 binding arm']).upper() or (q2[:-20] if len(q2) > 20 else '')
            P.append([s(p['Pair #']), s(p['P1 standard name']), a1, s(p['P2 standard name']), a2, q1, q2,
                      s(p.get('Genome match (to set LOC)')) or s(p.get('Template match')), s(p.get('Pair flag')),
                      s(p.get('P1 original name(s)')), s(p.get('P2 original name(s)'))])
        short, long_, aka = names_for(gene)
        gv = r.get('# pairs verified in genome')
        S.append(dict(id=f'{by}:{sid}', sid=sid, by=by, file=os.path.basename(f), gene=gene, full=s(r.get('Gene / protein name')),
            org=s(r.get('Organism')) or 'Not stated', loc=s(r.get('LOC / GeneID')), gloc=s(r.get('Genome LOC (genome_cds.fa match)')),
            short=short, long=long_, aka=aka, amp=s(r['Amplifier']), fl=s(r.get('Fluorophore')), gv=gv if gv not in (None, '') else None,
            checked=bool(s(r.get('LOC source')) or s(r.get('Genome LOC (genome_cds.fa match)'))),
            weak=s(r.get('Genome confidence')) == 'Low' and s(r.get('LOC source')) == 'genome_cds.fa match',
            gname=s(r.get('NCBI protein name (genome_cds.fa)')), names=s(r.get('Names used in source files')),
            check=s(r.get('Check notes')), status=s(r.get('Status')), tube=s(r.get('Tube / location')), notes=s(r.get('Notes')), p=P))

# ---------- probe sets submitted with templates/HCR_probe_submission.xlsx and dropped into user_submit/ ----------
# Only typed-in cells are read (never formula results), so a file saved by Excel, Google Sheets or LibreOffice works.
# Every pair is checked here; a pair that fails is left out and reported in the build log, never breaking the page.
DEFAULT_AMPS = {'B1': ('AF546', 'GAGGAGGGCAGCAAACGG', 'AA', 'TA', 'GAAGAGTCTTCCTTTACG'),
                'B2': ('AF647', 'CCTCGTAAATCCTCATCA', 'AA', 'AA', 'ATCATCCAGTAAACCGCC'),
                'B3': ('AF488', 'GTCCCTGCCTCTATATCT', 'TT', 'TT', 'CCACTCAACTTTAACCCG'),
                'B4': ('AF514', 'CCTCAACCTACCTCCAAC', 'AA', 'AT', 'TCTCACCATATTCGCTTC'),
                'B5': ('AF594', 'CTCACTCCCAATCTCTAT', 'AA', 'AA', 'CTACCCTACAAATCCAAT')}
amps = amps or {}
for a, (fl, p1, s1, s2, p2) in DEFAULT_AMPS.items():
    amps.setdefault(a, dict(fl=fl, p1=p1, s1=s1, s2=s2, p2=p2))
subdir = sys.argv[3] if len(sys.argv) > 3 else os.path.join(here, 'user_submit')
sub_files = sorted(f for f in glob.glob(os.path.join(subdir, '*.xlsx')) if not os.path.basename(f).startswith(('_', '~$')))
def clean_seq(v): return re.sub(r'[\s\-]', '', s(v)).upper().replace('U', 'T')
def amp_of(q1, q2):
    a1 = next((a for a, A in amps.items() if q1.startswith(A['p1'])), None)
    a2 = next((a for a, A in amps.items() if q2.endswith(A['p2'])), None)
    return a1, a2
def gene_key(v):
    v = s(v).replace(' ', '')
    if re.fullmatch(r'\d+', v): return 'GeneID:' + v
    m = re.fullmatch(r'(?i)(?:NCBI)?GeneID:?(\d+)', v)
    return 'GeneID:' + m.group(1) if m else v
def as_date(v):
    if isinstance(v, (datetime.date, datetime.datetime)): return v.strftime('%Y-%m-%d')
    return s(v)
used_prefix = {}                                   # set-ID prefix -> designer, so two people never share one
for x in S: used_prefix.setdefault(re.sub(r'[-\d]+$', '', x['sid']).rstrip('-') or x['sid'], x['by'])
subs = []
for f in sub_files:
    base = os.path.basename(f)
    try:
        wb = load_workbook(f, data_only=True)
        ab, gs, ps = wb['1 About you'], wb['2 Genes'], wb['3 Probes']
    except Exception as e:
        print(f'user_submit/{base}: skipped, not a submission form ({e})'); continue
    info = {s(ab.cell(r, 1).value): ab.cell(r, 2).value for r in range(1, 20) if s(ab.cell(r, 1).value)}
    who = s(info.get('Your name'))
    if not who: print(f'user_submit/{base}: skipped, "Your name" is empty'); continue
    subs.append((as_date(info.get('Date of this submission')), base, who, info, gs, ps))
subs.sort(key=lambda t: (t[2].lower(), t[0], t[1]))      # each person's files in date order, so set numbers stay put
counter = Counter(); n_sub = 0
for date, base, who, info, gs, ps in subs:
    init = re.sub(r'[^A-Za-z]', '', s(info.get('Initials for your set IDs'))).upper()[:4] or ''.join(w[0] for w in who.split()).upper()[:4]
    pre = init; k = 2
    while used_prefix.get(pre, who) != who: pre = f'{init}{k}'; k += 1
    used_prefix[pre] = who
    if who not in designers: designers.append(who)
    genes_in = {}
    hdr = {re.sub(r'\s*\*$', '', s(gs.cell(5, j).value)).lower(): j for j in range(1, 40) if s(gs.cell(5, j).value)}
    def col(r, name): j = hdr.get(name); return gs.cell(r, j).value if j else None
    for r in range(6, 6 + 30):
        num = col(r, 'gene #'); short = s(col(r, 'gene short name'))
        if not short: continue
        genes_in[int(num) if isinstance(num, (int, float)) else r - 5] = dict(
            short=short, long=s(col(r, 'gene long name')), org=re.sub(r'^Other:.*$', '', s(col(r, 'species'))) or 'Not stated',
            loc=gene_key(col(r, 'gene id or loc')), tx=s(col(r, 'transcript targeted')), amp=s(col(r, 'amplifier')).upper(),
            dye=s(col(r, 'dye you used')), region=s(col(r, 'where the probes bind')), method=s(col(r, 'how you designed them')),
            works=s(col(r, 'did it work?')), tissue=s(col(r, 'tissue / stage tested')), odate=as_date(col(r, 'order date')),
            vendor=s(col(r, 'vendor / order id')), store=s(col(r, 'where the tubes are stored')), notes=s(col(r, 'notes')))
    pairs = defaultdict(list); problems = []
    for r in range(5, ps.max_row + 1):
        gn, q1, q2 = ps.cell(r, 1).value, clean_seq(ps.cell(r, 4).value), clean_seq(ps.cell(r, 5).value)
        if gn in (None, '') and not q1 and not q2: continue
        try: gn = int(gn)
        except (TypeError, ValueError): problems.append(f'row {r}: no gene number'); continue
        if gn not in genes_in: problems.append(f'row {r}: gene #{gn} is not filled in on "2 Genes"'); continue
        if not q1 or not q2: problems.append(f'row {r}: P1 or P2 missing'); continue
        if re.search(r'[^ACGT]', q1 + q2): problems.append(f'row {r}: letters other than A, C, G, T'); continue
        a1, a2 = amp_of(q1, q2)
        if not a1 or not a2 or a1 != a2: problems.append(f'row {r}: initiators not recognized or mixed ({a1 or "?"} / {a2 or "?"})'); continue
        pairs[gn].append((q1, q2, a1, s(ps.cell(r, 7).value), s(ps.cell(r, 8).value)))
    for gn in sorted(genes_in):
        g = genes_in[gn]; P0 = pairs.get(gn, [])
        if not P0: problems.append(f'gene #{gn} ({g["short"]}): no usable pairs, left out'); continue
        seen_amp = Counter(p[2] for p in P0).most_common(1)[0][0]
        note = []
        if g['amp'] in amps and g['amp'] != seen_amp:
            note.append(f'The form says {g["amp"]}, but the sequences carry the {seen_amp} initiator; shown as {seen_amp}.')
        amp = seen_amp
        if any(p[2] != amp for p in P0):
            problems.append(f'gene #{gn} ({g["short"]}): {sum(p[2] != amp for p in P0)} pair(s) on another amplifier left out')
            P0 = [p for p in P0 if p[2] == amp]
        counter[pre] += 1; sid = f'{pre}-{counter[pre]:03d}'
        label = re.sub(r'[^A-Za-z0-9.\-()]+', '-', g['short'])
        P = []
        for i, (q1, q2, _, o1, o2) in enumerate(P0, 1):
            nn = f'{i:02d}'
            P.append([str(i), f'{label}_{amp}_P1_{nn}', q1[20:], f'{label}_{amp}_P2_{nn}', q2[:-20], q1, q2, '',
                      '' if len(q1) == 45 and len(q2) == 45 else f'lengths {len(q1)} and {len(q2)} nt', o1, o2])
        short, long_, aka = names_for(g['short'])
        x = dict(id=f'{who}:{sid}', sid=sid, by=who, file='user_submit/' + base, gene=g['short'], full=g['long'] or long_,
                 org=g['org'], loc=g['loc'], gloc='', short=g['short'], long=g['long'] or long_, aka=aka, amp=amp,
                 fl=g['dye'] if g['dye'] and not g['dye'].startswith('Other') else amps[amp]['fl'], gv=None, checked=False, weak=False,
                 gname='', names=g['short'], check=' '.join(note), status=g['works'], tube=g['store'], notes=g['notes'], p=P,
                 sub=dict(lab=s(info.get('Lab / PI')), inst=s(info.get('Institution')), date=date, tx=g['tx'], region=g['region'],
                          method=g['method'], tissue=g['tissue'], odate=g['odate'], vendor=g['vendor']))
        S.append(x); n_sub += 1
    print(f'user_submit/{base}: {who} ({pre}), {(lambda n: f"{n} set" + ("s" if n != 1 else ""))(sum(1 for x in S if x["file"] == "user_submit/" + base))}'
          + (''.join(f'\n   - {p}' for p in problems) if problems else ', no problems'))

# oligos that do not carry the initiator of their set's amplifier
for x in S:
    A = amps.get(x['amp'])
    if not A: continue
    bad = [p[0] for p in x['p'] if (p[5] and not p[5].startswith(A['p1'])) or (p[6] and not p[6].endswith(A['p2']))]
    if bad:
        x['check'] = '; '.join(filter(None, [x['check'], f"Pair(s) {', '.join(bad)} do not carry the {x['amp']} initiator; check the amplifier"]))

# sets that share oligos with another designer's set
owner = defaultdict(set)
for x in S:
    for p in x['p']:
        for q in (p[5], p[6]):
            if q: owner[q].add(x['id'])
byid = {x['id']: x for x in S}
for x in S:
    shared = Counter(o for p in x['p'] for q in (p[5], p[6]) if q for o in owner[q] if o != x['id'] and byid[o]['by'] != x['by'])
    for o, n in shared.items():
        y = byid[o]
        x['check'] = '; '.join(filter(None, [x['check'], f"Shares {n} oligos with {y['by']}'s set {y['sid']} ({y['gene']})"]))

# group sets into genes: by genome locus when there is one, otherwise by recorded LOC, otherwise by name + organism
def key(x):
    # a weak genome match (under half the pairs, LOC not recorded by the designer) does not decide the gene
    if x['weak']: return x['org'] + '|' + x['gene'].lower()
    l = re.findall(r'LOC\d+|GeneID:\d+', x['gloc'] or x['loc'])
    return l[0] if l else (x['org'] + '|' + x['gene'].lower())
G = defaultdict(list)
for x in S: G[key(x)].append(x['id'])
# a set without any LOC joins an existing gene with the same name and organism
named = {}
for k, ids in G.items():
    if k.startswith(('LOC', 'GeneID')):
        for i in ids: named.setdefault((byid[i]['org'], byid[i]['gene'].lower()), k)
for k in [k for k in G if not k.startswith(('LOC', 'GeneID'))]:
    org, name = k.split('|', 1)
    if (org, name) in named:
        G[named[(org, name)]] += G.pop(k)
genes = []
for k, ids in G.items():
    xs = [byid[i] for i in ids]
    label = Counter(x['short'] or x['gene'] for x in xs).most_common(1)[0][0]
    long_ = next((x['long'] for x in xs if x['long']), '')
    aka = sorted({a.strip() for x in xs for a in x['aka'].split(',') if a.strip()})
    genes.append(dict(k=k, label=label, long=long_, aka=aka, names=sorted({x['gene'] for x in xs}), org=xs[0]['org'],
                      prot=next((x['gname'] or x['full'] for x in xs if (x['gname'] or x['full'])), ''),
                      loc=k if k.startswith(('LOC', 'GeneID')) else '', sets=ids))
genes.sort(key=lambda g: g['label'].lower())
# hairpins (B1-B20) and transcriptome-wide designs, if present next to the output page
hair = {}
hpath = os.path.join(here, 'hairpins.xlsx')
if os.path.exists(hpath):
    for r in table(load_workbook(hpath, data_only=True, read_only=True)['Hairpins']):
        a = s(r.get('Amplifier'))
        if re.fullmatch(r'B\d+', a):
            hair[a] = dict(dye=s(r.get('Dye in source document')), I1=s(r.get('I1')), I2=s(r.get('I2')), H1=s(r.get('H1 (3′ dye)')), H2=s(r.get('H2 (5′ dye)')))
outdir = os.path.dirname(os.path.abspath(out))
SPECIES = (('bany', 'Bicyclus anynana'), ('danio', 'Danio rerio'), ('nfur', 'Nothobranchius furzeri (killifish)'),
           ('dmel', 'Drosophila melanogaster'), ('cele', 'Caenorhabditis elegans'), ('mmus', 'Mus musculus'), ('hsap', 'Homo sapiens'))
tx = [dict(id=k, name=v) for k, v in SPECIES
      if os.path.exists(os.path.join(outdir, 'tx', k, 'index.js')) or os.path.exists(os.path.join(here, 'tx', k, 'index.js'))]
import base64
guide = {}
gdir = os.path.join(here, 'guide')
if os.path.isdir(gdir):
    for f in sorted(os.listdir(gdir)):
        if f.endswith('.jpg'):
            guide[f[:-4]] = 'data:image/jpeg;base64,' + base64.b64encode(open(os.path.join(gdir, f), 'rb').read()).decode()
data = dict(amps=amps, sets=byid, genes=genes, designers=designers, hairpins=hair, tx=tx, guide=guide,
            hairpinNote='Hairpin sequences from B1_B20.docx; the dye in that document is the one it lists, not necessarily what the lab orders.', built=datetime.date.today().isoformat(),
            source=', '.join(os.path.basename(f) for f in files) + (f' and {len(subs)} submitted file{"s" if len(subs) != 1 else ""}' if subs else ''))
open(out, 'w', encoding='utf-8').write(TEMPLATE.replace('/*__DATA__*/null', json.dumps(data, separators=(',', ':'))))
print(f'wrote {out}: {len(files)} files, {len(genes)} genes, {len(S)} sets, {sum(len(x["p"]) for x in S)} pairs')
