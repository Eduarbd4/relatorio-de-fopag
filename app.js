const ACCESS_PASSWORD = 'DANE-SE2015';
const $ = (id) => document.getElementById(id);
const state = { workbook:null, sheetName:'', employees:[], filtered:[], current:null, printMode:'one', editedReports:{} };
const normalize = (v) => String(v ?? '').normalize('NFD').replace(/[\u0300-\u036f]/g,'').trim().toUpperCase();
const num = (v) => { if(typeof v==='number') return Number.isFinite(v)?v:0; const n=Number(String(v??'').replace(/[^0-9,.-]/g,'').replace(/\./g,'').replace(',','.')); return Number.isFinite(n)?n:0; };
const money = (v) => typeof v==='string' && /sem comiss/i.test(v) ? 'Sem comissão' : num(v).toLocaleString('pt-BR',{style:'currency',currency:'BRL'});
const cleanUnit = (v) => String(v??'').replace(/\s+/g,' ').trim();
const safe = (v) => String(v??'').replace(/[&<>'"]/g,c=>({'&':'&amp;','<':'&lt;','>':'&gt;',"'":'&#39;','"':'&quot;'}[c]));
const isMeaningful = (v) => v!==null && v!==undefined && String(v).trim()!=='' && !(typeof v==='number' && Math.abs(v)<0.000001) && !['0','0.0','-'].includes(String(v).trim());

function unlock(){ sessionStorage.setItem('daneAuth','1'); $('login').classList.add('hidden'); $('app').classList.remove('hidden'); }
if(sessionStorage.getItem('daneAuth')==='1') unlock();
$('loginForm').addEventListener('submit',e=>{e.preventDefault();if($('password').value===ACCESS_PASSWORD){$('loginError').textContent='';unlock()}else{$('loginError').textContent='Senha incorreta. Tente novamente.';$('password').select()}});
$('togglePassword').onclick=()=>{$('password').type=$('password').type==='password'?'text':'password'};
$('logout').onclick=()=>{sessionStorage.removeItem('daneAuth');location.reload()};

const input=$('fileInput'), drop=$('dropzone');
input.onchange=()=>input.files[0]&&loadFile(input.files[0]);
['dragenter','dragover'].forEach(ev=>drop.addEventListener(ev,e=>{e.preventDefault();drop.classList.add('drag')}));
['dragleave','drop'].forEach(ev=>drop.addEventListener(ev,e=>{e.preventDefault();drop.classList.remove('drag')}));
drop.addEventListener('drop',e=>{const f=e.dataTransfer.files[0];if(f)loadFile(f)});

async function loadFile(file){
  if(!/\.(xlsx|xlsm?|xls)$/i.test(file.name)){toast('Selecione um arquivo Excel válido.');return}
  if(typeof XLSX==='undefined'){toast('Não foi possível carregar o leitor de Excel. Verifique a internet.');return}
  try{
    drop.querySelector('strong').textContent='Lendo planilha…';
    const buffer=await file.arrayBuffer(); state.workbook=XLSX.read(buffer,{type:'array',cellDates:true}); state.editedReports={};
    const closing=state.workbook.SheetNames.filter(n=>/fechamento/i.test(n));
    if(!closing.length) throw new Error('Nenhuma aba com “Fechamento” foi encontrada.');
    $('sheetSelect').innerHTML=closing.map(n=>`<option value="${safe(n)}">${safe(n)}</option>`).join('');
    state.sheetName=chooseLatestSheet(closing); $('sheetSelect').value=state.sheetName;
    $('fileInfo').textContent=`${file.name} • ${(file.size/1024/1024).toFixed(1)} MB`; $('fileInfo').classList.remove('hidden');
    $('sheetControls').classList.remove('hidden'); $('exportControls').classList.remove('hidden');
    processSheet();
  }catch(err){toast(err.message||'Não foi possível ler a planilha.');console.error(err)}finally{drop.querySelector('strong').textContent='Selecionar outra planilha'}
}
function chooseLatestSheet(names){
  const months={jan:1,fev:2,mar:3,abr:4,mai:5,jun:6,jul:7,ago:8,set:9,out:10,nov:11,dez:12};
  return [...names].sort((a,b)=>{const score=n=>{const x=normalize(n).toLowerCase();let m=0;for(const [k,v] of Object.entries(months))if(x.includes(k))m=v;const y=+(x.match(/(?:20)?(\d{2})/)?.[1]||0);return y*20+m+(x.includes('fechado')||x.includes('fechada')?0.5:0)};return score(b)-score(a)})[0];
}
$('sheetSelect').onchange=()=>{state.sheetName=$('sheetSelect').value;processSheet()};
$('competence').oninput=()=>renderEmployees();
$('clearFile').onclick=()=>{state.workbook=null;state.employees=[];input.value='';$('fileInfo').classList.add('hidden');$('sheetControls').classList.add('hidden');$('exportControls').classList.add('hidden');$('results').classList.add('hidden');$('emptyState').classList.remove('hidden');drop.querySelector('strong').textContent='Selecionar planilha'};

function parseEmployees(rows){
  const first=rows.findIndex(r=>['FUNCIONARIO','COLABORADOR'].includes(normalize(r?.[0]))&&normalize(r?.[1])==='UNIDADE');
  if(first<0)throw new Error('Cabeçalho de colaboradores não encontrado.');
  const main=rows[first], modern=main.some(h=>normalize(h)==='VALOR TOTAL VT');
  const summaryCol=main.findIndex(h=>normalize(h)==='NOME');
  let headers=main;const employees=[];
  const label=h=>String(h||'').replace(/\s+/g,' ').trim();
  for(let r=first+1;r<rows.length;r++){
    const row=rows[r]||[],name=String(row[0]??'').trim(),unit=cleanUnit(row[1]);
    if(['FUNCIONARIO','COLABORADOR'].includes(normalize(name))&&normalize(unit)==='UNIDADE'){headers=row;continue;}
    if(!name||!unit||normalize(name)===normalize(unit)||/^(MATRIZ|FILIAL|CENTRAL)(\s|-|$)/.test(normalize(name)))continue;
    const at=(re)=>headers.findIndex(h=>re.test(normalize(h)));
    const get=(re)=>{const c=at(re);return c<0?null:row[c]};
    const bonusItems=[];
    const start=modern?17:12,end=modern?35:26;
    for(let c=start;c<=end;c++){
      if(modern&&[29,30].includes(c))continue;
      if(isMeaningful(row[c])&&typeof row[c]==='number')bonusItems.push({label:label(headers[c]||((c>=12&&c<=15)?`Semana ${c-11}`:`Bonificação ${c}`)),value:row[c]});
    }
    const leader=main.findIndex(h=>normalize(h)==='PONTUACAO CHECKLIST')-2;
    const lr=leader>=0?rows.find(x=>normalize(x?.[leader])===normalize(name)&&normalize(x?.[leader+1])===normalize(unit)):null;
    if(lr)for(let c=leader+4;c<main.length;c++){
      if(normalize(main[c])==='BONUS TOTAL')break;
      if(main[c]&&typeof lr[c]==='number'&&isMeaningful(lr[c]))bonusItems.push({label:'Liderança — '+label(main[c]),value:lr[c]});
    }
    const e={id:employees.length,name,unit,bonusItems,
      vaValue:get(/^(VALOR VA|VALOR DIARIO VA)$/),vaDays:get(/^(QTD VA|QUANTIDADE TOTAL VA)$/),vaTotal:get(/^(TOTAL VA|VALOR TOTAL VA)$/),
      vtValue:get(/^(VALOR VT|VALOR DIARIO VT)$/),vtDays:get(/^(QTD VT|QUANTIDADE TOTAL VT)$/),vtTotal:get(/^(TOTAL VT|VALOR TOTAL VT)$/),
      vaObs:modern?row[15]:row[8],vtObs:row[8],vaFormat:modern?row[14]:null,vtFormat:modern?row[7]:null,
      vaExtra:modern?row[10]:null,vtExtra:modern?row[3]:null,
      benefitObs:modern?null:row[8],hours:get(/HORA EXTRA/),obs:modern?null:row[10],bonusObs:modern?row[29]:row[28],
      bonusTotal:get(/^BONUS TOTAL$/),att:get(/^ATESTADO/),absence:get(/^FALTA/),advance:get(/ADIANTAMENTO/),
      withdrawal:get(/^(RETIRADA|B:.*RETIRADAS)/),caju:modern?row[41]:null,otherDiscount:modern?row[42]:null,
      deductTotal:get(/^RETIRADA TOTAL$/),deductDesc:get(/^DESCRICAO$/),commission:get(/^COMISSAO/),sales:get(/^TOTAL REALIZADO$/)};
    const summary=summaryCol<0?null:rows.find(x=>normalize(x?.[summaryCol])===normalize(name));
    if(summary){e.commission=summary[summaryCol+1]??e.commission;e.bonusTotal=summary[summaryCol+2]??e.bonusTotal;e.summaryObs=summary[summaryCol+6];}
    else if(lr){const total=main.findIndex((h,c)=>c>leader&&normalize(h)==='BONUS TOTAL');e.bonusTotal=num(e.bonusTotal)+num(lr[total]);}
    e.notes=[e.obs,e.bonusObs,e.deductDesc,e.summaryObs].filter(isMeaningful).join(' | ');
    employees.push(e);
  }
  return consolidateEmployees(employees);
}
function employeeIdentity(name){
  return normalize(name).replace(/\s*\((SG|VR)\)\s*/g,' ').replace(/\s+(FLG|BSB|LGO|LAGO|MNE|PKS|STE|GO|IGT|CD)$/,'').replace(/\s+/g,' ').trim();
}
function consolidateEmployees(employees){
  const groups=new Map();
  for(const e of employees){const key=employeeIdentity(e.name);if(!groups.has(key))groups.set(key,[]);groups.get(key).push(e);}
  return [...groups.values()].map((sources,id)=>{
    if(sources.length===1)return {...sources[0],id,sources};
    const combined={...sources[0],id,name:employeeIdentity(sources[0].name),unit:[...new Set(sources.map(e=>e.unit))].join(' / '),sources};
    for(const key of ['vaDays','vaTotal','vtDays','vtTotal','bonusTotal','commission','sales','advance','withdrawal','caju','otherDiscount','deductTotal','att','absence'])combined[key]=sources.reduce((t,e)=>t+num(e[key]),0);
    combined.bonusItems=sources.flatMap(e=>e.bonusItems.map(x=>({...x,label:e.unit+' — '+x.label})));
    combined.notes=sources.map(e=>e.unit+': '+(e.notes||'Sem observações adicionais')).join(' | ');
    return combined;
  });
}

function processSheet(){
  const ws=state.workbook.Sheets[state.sheetName];
  const rows=XLSX.utils.sheet_to_json(ws,{header:1,defval:null,raw:true});
  const employees=parseEmployees(rows);
  state.employees=employees; $('competence').value=inferCompetence(state.sheetName); $('employeeCount').textContent=employees.length;
  $('emptyState').classList.add('hidden');$('results').classList.remove('hidden');
  const warning=employees.length?'': 'Nenhum funcionário foi identificado nesta aba. Confira se ela segue o modelo de fechamento.';
  $('notice').textContent=warning;$('notice').classList.toggle('hidden',!warning);renderEmployees();
}
function inferCompetence(name){const map={JAN:'Janeiro',FEV:'Fevereiro',MAR:'Março',ABR:'Abril',MAI:'Maio',JUN:'Junho',JUL:'Julho',AGO:'Agosto',SET:'Setembro',OUT:'Outubro',NOV:'Novembro',DEZ:'Dezembro'};const n=normalize(name);const key=Object.keys(map).find(k=>n.includes(k));const yy=n.match(/(?:^|\D)(20)?(\d{2})(?:\D|$)/)?.[2];return key&&yy?`${map[key]}/20${yy}`:(key?map[key]:'')}
function renderEmployees(){const q=normalize($('search').value);state.filtered=state.employees.filter(e=>!q||normalize(e.name+' '+e.unit).includes(q));$('employeeList').innerHTML=state.filtered.length?state.filtered.map(e=>`<article class="employee-card"><div class="avatar">${safe(e.name[0])}</div><div class="employee-meta"><strong>${safe(e.name)}</strong><small>${safe(e.unit)}</small></div><button data-id="${e.id}" aria-label="Ver relatório de ${safe(e.name)}">Ver</button></article>`).join(''):'<div class="no-data">Nenhum funcionário encontrado.</div>';document.querySelectorAll('[data-id]').forEach(b=>b.onclick=()=>openReport(state.employees[+b.dataset.id]))}
$('search').oninput=renderEmployees;

function reportHTML(e){
  const commission=typeof e.commission==='string'&&/sem comiss/i.test(e.commission)?0:num(e.commission),bonus=num(e.bonusTotal),discount=num(e.deductTotal),result=commission+bonus-discount;
  const sources=e.sources||[e],benefits=[];
  for(const source of sources){
    for(const [key,title] of [['va','Vale-alimentação'],['vt','Vale-transporte']]){
      if(isMeaningful(source[key+'Total']))benefits.push([title+(sources.length>1?' — '+source.unit:''),source[key+'Days'],money(source[key+'Value']),money(source[key+'Total']),[source[key+'Format'],source[key+'Obs'],isMeaningful(source[key+'Extra'])?'Extras: '+source[key+'Extra']:null].filter(isMeaningful).join(' | ')||'Conforme fechamento']);
    }
  }
  const bonuses=e.bonusItems.filter(x=>isMeaningful(x.value));if(!bonuses.length&&bonus)bonuses.push({label:'Bonificações e prêmios',value:bonus});
  const discounts=[];
  for(const source of sources){
    const prefix=sources.length>1?source.unit+' — ':'';
    for(const [key,title,kind] of [['att','Atestado informado','days'],['absence','Faltas','days'],['advance','Adiantamento salarial','money'],['withdrawal','Retiradas de produtos','money'],['caju','Caju — benefícios','money'],['otherDiscount','Outros descontos','money'],['hours','Horas extras/noturnas — ocorrência informada','text']]){
      if(isMeaningful(source[key]))discounts.push([prefix+title,kind==='money'?money(source[key]):kind==='days'?source[key]+' dia(s)':String(source[key])]);
    }
  }
  if(!discounts.length)discounts.push(['Sem descontos ou ocorrências informados','—']);
  const detailedBonus=bonuses.reduce((t,x)=>t+num(x.value),0);if(Math.abs(bonus-detailedBonus)>0.01)bonuses.push({label:'Diferença entre o total do fechamento e os itens detalhados — conferir',value:bonus-detailedBonus});
  const competence=$('competence').value||inferCompetence(state.sheetName);
  return `<article class="report-page"><header class="report-header"><div class="report-logo">dane-se<sup>®</sup></div><div class="report-title"><h1>RELATÓRIO DE DETALHAMENTO DE<br>PAGAMENTO</h1><p>Documento complementar ao contracheque | Competência: ${safe(competence)}</p></div></header><section class="identity"><div><span>COLABORADOR(A)</span><strong>${safe(e.name)}</strong></div><div><span>UNIDADE</span><strong>${safe(e.unit)}</strong></div><div><span>COMPETÊNCIA</span><strong>${safe(competence)}</strong></div></section><section class="totals"><div class="total-box"><span>COMISSÃO</span><strong>${money(e.commission)}</strong></div><div class="total-box"><span>BONIFICAÇÕES</span><strong>${money(bonus)}</strong></div><div class="total-box"><span>DESCONTOS/RETIRADAS</span><strong>${money(discount)}</strong></div><div class="total-box result"><span>RESULTADO VARIÁVEL</span><strong>${money(result)}</strong></div></section><p class="formula-note">Resultado variável = comissão + bonificações - descontos/retiradas informados. VA e VT são demonstrados separadamente. Valores conforme fechamento; não representam o salário líquido completo.</p><section class="report-section"><h2>RESUMO POR UNIDADE</h2><table class="report-table"><thead><tr><th>UNIDADE</th><th>VA</th><th>VT</th><th>BÔNUS</th><th>DESCONTOS</th></tr></thead><tbody>${sources.map(source=>`<tr><td>${safe(source.unit)}</td><td>${money(source.vaTotal)}</td><td>${money(source.vtTotal)}</td><td>${money(source.bonusTotal)}</td><td>${money(source.deductTotal)}</td></tr>`).join('')}</tbody></table></section><section class="report-section"><h2>COMISSÃO</h2><table class="report-table"><thead><tr><th>UNIDADE</th><th class="right">VENDAS REALIZADAS</th><th class="right">COMISSÃO</th></tr></thead><tbody>${sources.map(source=>`<tr><td>${safe(source.unit)}</td><td class="right">${money(source.sales)}</td><td class="right">${money(source.commission)}</td></tr>`).join('')}<tr class="total-row"><td>Total</td><td class="right">${money(e.sales)}</td><td class="right">${money(e.commission)}</td></tr></tbody></table></section><section class="report-section"><h2>BENEFÍCIOS</h2><table class="report-table"><thead><tr><th>BENEFÍCIO</th><th>DIAS</th><th>VALOR DIÁRIO</th><th>TOTAL</th><th>OBSERVAÇÃO</th></tr></thead><tbody>${benefits.length?benefits.map(x=>`<tr>${x.map((v,i)=>`<td class="${i>1&&i<4?'right':''}">${safe(v)}</td>`).join('')}</tr>`).join(''):'<tr><td colspan="5">Sem benefícios informados</td></tr>'}</tbody></table></section><section class="report-section"><h2>BONIFICAÇÕES E PRÊMIOS</h2><table class="report-table"><thead><tr><th>BONIFICAÇÃO / PRÊMIO</th><th class="right">VALOR</th></tr></thead><tbody>${bonuses.map(x=>`<tr><td>${safe(x.label)}</td><td class="right">${money(x.value)}</td></tr>`).join('')}<tr class="total-row"><td>Total de bonificações</td><td class="right">${money(bonus)}</td></tr></tbody></table></section><section class="report-section"><h2>DESCONTOS, RETIRADAS E OCORRÊNCIAS</h2><table class="report-table"><thead><tr><th>DESCONTO / RETIRADA</th><th class="right">QUANTIDADE OU VALOR</th></tr></thead><tbody>${discounts.map(x=>`<tr><td>${safe(x[0])}</td><td class="right">${safe(x[1])}</td></tr>`).join('')}<tr class="total-row"><td>Total de descontos/retiradas informado</td><td class="right">${money(discount)}</td></tr></tbody></table></section><div class="details-note">Detalhamento: ${safe(e.notes||'Conforme informações registradas no fechamento.')}</div><footer class="report-footer"><span>Documento complementar ao contracheque. Não substitui o demonstrativo oficial de pagamento.</span><span>DANE-SE | ${safe(competence)}</span></footer></article>`;
}
function scaleModal(){if(innerWidth<820){const s=Math.min((innerWidth-16)/794,1);$('reportPreview').style.width='794px';$('reportPreview').style.transform=`scale(${s})`;$('reportPreview').style.transformOrigin='top left';$('reportPreview').parentElement.style.width=`${794*s}px`}else{$('reportPreview').removeAttribute('style');$('reportPreview').parentElement.style.width='min(900px,100%)'}}
const reportKey=e=>`${state.sheetName}::${normalize(e.name)}::${normalize(e.unit)}::${$('competence').value}`;
function editableReport(e){return state.editedReports[reportKey(e)]||reportHTML(e)}
function saveCurrentEdit(){if(!state.current)return;const page=$('reportPreview').querySelector('.report-page');if(page)state.editedReports[reportKey(state.current)]=page.outerHTML}
function openReport(e){state.current=e;state.printMode='one';$('reportPreview').innerHTML=editableReport(e);const page=$('reportPreview').querySelector('.report-page');page.contentEditable='true';page.classList.add('editable-report');page.setAttribute('aria-label','Relatório editável de '+e.name);$('resetReport').classList.remove('hidden');$('reportModal').classList.remove('hidden');document.body.style.overflow='hidden';scaleModal()}
$('reportPreview').addEventListener('input',saveCurrentEdit);
$('resetReport').onclick=()=>{if(!state.current)return;delete state.editedReports[reportKey(state.current)];openReport(state.current);toast('Relatório restaurado com os dados da planilha.')};
$('closeModal').onclick=()=>{$('reportModal').classList.add('hidden');document.body.style.overflow=''};
$('reportModal').onclick=e=>{if(e.target===$('reportModal'))$('closeModal').click()};
$('printOne').onclick=()=>{state.printMode='one';window.print()};
$('printAll').onclick=()=>{state.printMode='all';state.current=null;$('reportPreview').innerHTML=state.employees.map(editableReport).join('');$('reportPreview').querySelectorAll('.report-page').forEach(p=>{p.contentEditable='false';p.classList.remove('editable-report')});$('resetReport').classList.add('hidden');$('reportModal').classList.remove('hidden');document.body.style.overflow='hidden';scaleModal();setTimeout(()=>window.print(),150)};
addEventListener('resize',()=>{if(!$('reportModal').classList.contains('hidden'))scaleModal()});
function toast(msg){$('toast').textContent=msg;$('toast').classList.remove('hidden');setTimeout(()=>$('toast').classList.add('hidden'),3500)}
