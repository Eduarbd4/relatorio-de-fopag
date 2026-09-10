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
    const buffer=await file.arrayBuffer(); state.workbook=XLSX.read(buffer,{type:'array',cellDates:true});
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
  return [...names].sort((a,b)=>{const score=n=>{const x=normalize(n).toLowerCase();let m=0;for(const [k,v] of Object.entries(months))if(x.includes(k))m=v;const y=+(x.match(/(?:20)?(\d{2})/)?.[1]||0);return y*20+m+(x.includes('em aberto')?100000:0)};return score(b)-score(a)})[0];
}
$('sheetSelect').onchange=()=>{state.sheetName=$('sheetSelect').value;processSheet()};
$('competence').oninput=()=>renderEmployees();
$('clearFile').onclick=()=>{state.workbook=null;state.employees=[];input.value='';$('fileInfo').classList.add('hidden');$('sheetControls').classList.add('hidden');$('exportControls').classList.add('hidden');$('results').classList.add('hidden');$('emptyState').classList.remove('hidden');drop.querySelector('strong').textContent='Selecionar planilha'};

function processSheet(){
  const ws=state.workbook.Sheets[state.sheetName];
  const rows=XLSX.utils.sheet_to_json(ws,{header:1,defval:null,raw:true});
  const headers=(rows[1]||[]).map(normalize);
  const fallback={name:0,unit:1,vaValue:2,vaDays:3,vaTotal:4,vtValue:5,vtDays:6,vtTotal:7,benefitObs:8,hours:9,obs:10,eligible:11,bonusTotal:27,att:29,absence:30,advance:31,withdrawal:32,deductTotal:33,deductDesc:34,commission:50};
  const index=(text,fb)=>{const i=headers.findIndex(h=>h.includes(text));return i>=0?i:fb};
  const col={...fallback,name:index('FUNCIONARIO',0),unit:index('UNIDADE',1),commission:index('COMISSAO',50),bonusTotal:index('BONUS TOTAL',27),att:index('ATESTADO',29),absence:index('FALTA',30),advance:index('ADIANTAMENTO',31),withdrawal:index('RETIRADA',32),deductTotal:index('RETIRADA TOTAL',33),deductDesc:index('DESCRICAO',34)};
  const sectionHeaders=new Set(['VENDEDORES','VENDEDOR','SUBGERENCIA','GERENCIA','ESTOQUE','CAIXA','ADMINISTRATIVO','ADICIONAIS','V']);
  const employees=[];
  for(let r=2;r<rows.length;r++){
    const row=rows[r]||[], name=String(row[col.name]??'').trim(), unit=cleanUnit(row[col.unit]);
    if(!name||!unit||normalize(name)==='FUNCIONARIO'||normalize(unit)==='UNIDADE'||sectionHeaders.has(normalize(name))||normalize(name)===normalize(unit)||/^(MATRIZ|FILIAL|CENTRAL)(\s|-|$)/.test(normalize(name)))continue;
    const bonusItems=[];
    for(let c=12;c<=26;c++){if(isMeaningful(row[c]))bonusItems.push({label:String(rows[1]?.[c]||`Bonificação ${c-11}`).replace(/\s*\[[^\]]*\]/g,'').replace(/\n/g,' '),value:row[c]})}
    // Subgerência: detalhamento adicional do segundo bloco, quando houver.
    if(/\(SG\)/i.test(name)){for(let c=55;c<=67;c++){if(isMeaningful(row[c]))bonusItems.push({label:String(rows[1]?.[c]||'Bonificação').replace(/\s*\[[^\]]*\]/g,'').replace(/\n/g,' '),value:row[c]})}}
    const employee={id:employees.length,name,unit,vaValue:row[2],vaDays:row[3],vaTotal:row[4],vtValue:row[5],vtDays:row[6],vtTotal:row[7],benefitObs:row[8],hours:row[9],obs:row[10],bonusTotal:row[col.bonusTotal],bonusItems,att:row[col.att],absence:row[col.absence],advance:row[col.advance],withdrawal:row[col.withdrawal],deductTotal:row[col.deductTotal],deductDesc:row[col.deductDesc],commission:row[col.commission]};
    // Prefer the summary block for final commission/bonuses when it contains the same employee.
    const summary=rows.find(x=>normalize(x?.[70])===normalize(name));
    if(summary){employee.commission=summary[71]??employee.commission;employee.bonusTotal=summary[72]??employee.bonusTotal;employee.absence=summary[73]??employee.absence;employee.deductTotal=summary[74]??employee.deductTotal;employee.hours=summary[75]??employee.hours;employee.summaryObs=summary[76]}
    employees.push(employee);
  }
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
  const benefits=[];if(isMeaningful(e.vaTotal))benefits.push(['Vale-alimentação',e.vaDays,money(e.vaValue),money(e.vaTotal),e.benefitObs||'Conforme fechamento']);if(isMeaningful(e.vtTotal))benefits.push(['Vale-transporte',e.vtDays,money(e.vtValue),money(e.vtTotal),e.benefitObs||'Conforme fechamento']);
  const bonuses=e.bonusItems.filter(x=>isMeaningful(x.value));if(!bonuses.length&&bonus)bonuses.push({label:'Bonificações e prêmios',value:bonus});
  const discounts=[];if(isMeaningful(e.att))discounts.push(['Atestado informado',`${e.att} dia(s)`]);if(isMeaningful(e.absence))discounts.push(['Falta',`${e.absence} dia(s)`]);if(isMeaningful(e.advance))discounts.push(['Adiantamento salarial',money(e.advance)]);if(isMeaningful(e.withdrawal))discounts.push(['Retirada',money(e.withdrawal)]);if(isMeaningful(e.hours))discounts.push(['Hora extra/noturna',String(e.hours)]);if(!discounts.length)discounts.push(['Sem descontos ou ocorrências informados','—']);
  const competence=$('competence').value||inferCompetence(state.sheetName);
  return `<article class="report-page"><header class="report-header"><div class="report-logo">dane-se<sup>®</sup></div><div class="report-title"><h1>RELATÓRIO DE DETALHAMENTO DE<br>PAGAMENTO</h1><p>Documento complementar ao contracheque | Competência: ${safe(competence)}</p></div></header><section class="identity"><div><span>COLABORADOR(A)</span><strong>${safe(e.name)}</strong></div><div><span>UNIDADE</span><strong>${safe(e.unit)}</strong></div><div><span>COMPETÊNCIA</span><strong>${safe(competence)}</strong></div></section><section class="totals"><div class="total-box"><span>COMISSÃO</span><strong>${money(e.commission)}</strong></div><div class="total-box"><span>BONIFICAÇÕES</span><strong>${money(bonus)}</strong></div><div class="total-box"><span>DESCONTOS/RETIRADAS</span><strong>${money(discount)}</strong></div><div class="total-box result"><span>RESULTADO VARIÁVEL</span><strong>${money(result)}</strong></div></section><p class="formula-note">Resultado variável = comissão + bonificações - descontos/retiradas informados. VA e VT são demonstrados separadamente.</p><section class="report-section"><h2>BENEFÍCIOS</h2><table class="report-table"><thead><tr><th>BENEFÍCIO</th><th>DIAS</th><th>VALOR DIÁRIO</th><th>TOTAL</th><th>OBSERVAÇÃO</th></tr></thead><tbody>${benefits.length?benefits.map(x=>`<tr>${x.map((v,i)=>`<td class="${i>1&&i<4?'right':''}">${safe(v)}</td>`).join('')}</tr>`).join(''):'<tr><td colspan="5">Sem benefícios informados</td></tr>'}</tbody></table></section><section class="report-section"><h2>BONIFICAÇÕES E PRÊMIOS</h2><table class="report-table"><thead><tr><th>BONIFICAÇÃO / PRÊMIO</th><th class="right">VALOR</th></tr></thead><tbody>${bonuses.map(x=>`<tr><td>${safe(x.label)}</td><td class="right">${money(x.value)}</td></tr>`).join('')}<tr class="total-row"><td>Total de bonificações</td><td class="right">${money(bonus)}</td></tr></tbody></table></section><section class="report-section"><h2>DESCONTOS, RETIRADAS E OCORRÊNCIAS</h2><table class="report-table"><thead><tr><th>DESCONTO / RETIRADA</th><th class="right">QUANTIDADE OU VALOR</th></tr></thead><tbody>${discounts.map(x=>`<tr><td>${safe(x[0])}</td><td class="right">${safe(x[1])}</td></tr>`).join('')}<tr class="total-row"><td>Total descontado na folha</td><td class="right">${money(discount)}</td></tr></tbody></table></section><div class="details-note">Detalhamento: ${safe(e.deductDesc||e.obs||e.summaryObs||'Conforme informações registradas no fechamento.')}</div><footer class="report-footer"><span>Documento complementar ao contracheque. Não substitui o demonstrativo oficial de pagamento.</span><span>DANE-SE | ${safe(competence)}</span></footer></article>`;
}
function scaleModal(){if(innerWidth<820){const s=Math.min((innerWidth-16)/794,1);$('reportPreview').style.width='794px';$('reportPreview').style.transform=`scale(${s})`;$('reportPreview').style.transformOrigin='top left';$('reportPreview').parentElement.style.width=`${794*s}px`}else{$('reportPreview').removeAttribute('style');$('reportPreview').parentElement.style.width='min(900px,100%)'}}
const reportKey=e=>`${state.sheetName}::${normalize(e.name)}`;
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
