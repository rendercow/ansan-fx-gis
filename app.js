'use strict';
const $=id=>document.getElementById(id);
const state={companies:[],results:[],map:null,geocoder:null,markers:[],clusterer:null,selected:null,busy:false,coords:{},key:(window.FXGIS_CONFIG?.kakaoJavaScriptKey||localStorage.getItem('fxgis_kakao_js_key')||'')};
const norm=v=>v==null?'':String(v).trim();
const num=v=>{const n=Number(v);return Number.isFinite(n)?n:0};
const esc=s=>norm(s).replace(/[&<>"']/g,c=>({'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;',"'":'&#39;'}[c]));
const fmt=n=>Number(n).toLocaleString('ko-KR');
const label=v=>norm(v)||'—';
const setStatus=t=>$('status').textContent=t;
function readTable(wb,name){const sh=wb.Sheets[name];if(!sh)throw Error(`필수 시트가 없습니다: ${name}`);return sh.__rows||XLSX.utils.sheet_to_json(sh,{defval:'',raw:false});}
function ingest(wb){
 const bases=readTable(wb,'업체현황');const perf=readTable(wb,'연도별실적');const checks=readTable(wb,'점검이력');const risks=readTable(wb,'위험도평가');
 const pMap=new Map(),cMap=new Map(),rMap=new Map();
 for(const p of perf){const id=norm(p['업체ID']);if(!pMap.has(id))pMap.set(id,[]);pMap.get(id).push(p);}
 for(const c of checks){const id=norm(c['업체ID']);if(!cMap.has(id))cMap.set(id,[]);cMap.get(id).push(c);}
 for(const r of risks)rMap.set(norm(r['업체ID']),r);
 const seen=new Set();const entries=[];
 for(const r of bases){const id=norm(r['업체ID']);if(!id||seen.has(id))continue;seen.add(id);
  const district=norm(r['구']);if(!['상록구','단원구'].includes(district))continue;
  const address=norm(r['도로명주소(미검증)']);const risk=rMap.get(id)||{};
  const lat=Number(r['위도']),lng=Number(r['경도']);
  const validCoord=norm(r['위도'])!==''&&norm(r['경도'])!==''&&lat>=37.20&&lat<=37.45&&lng>=126.55&&lng<=127.05;
  const saved=state.coords[id];const cached=saved&&saved.address===address?saved:null;
  entries.push({id,name:norm(r['상호(가상)']),owner:norm(r['대표자(가상)']),district,address,registered:norm(r['등록일(가상)']),status:norm(r['영업상태']),r,perf:pMap.get(id)||[],checks:cMap.get(id)||[],risk:risk,grade:['고위험','주의','일반'].includes(norm(risk['위험등급']))?norm(risk['위험등급']):'미평가',score:num(risk['위험도합계']),lat:cached?cached.lat:validCoord?lat:null,lng:cached?cached.lng:validCoord?lng:null,geoStatus:cached?'확인완료':validCoord?'제공좌표':'미확인'});
 }
 state.companies=entries;state.selected=null;document.body.classList.remove('sheet-open');filter();renderDetail();setStatus(`업체 ${entries.length}개 로드 완료 (가상자료) · 좌표 미확인 ${entries.filter(c=>c.lat===null).length}건`);if(state.geocoder)geocodeAll();
}
function filter(){const s=norm($('search').value).toLowerCase(),d=$('district').value,g=$('risk').value;state.results=state.companies.filter(c=>(!d||c.district===d)&&(!g||(g==='좌표 미확인'?c.lat===null:c.grade===g))&&(!s||[c.name,c.owner,c.address,c.id].some(v=>v.toLowerCase().includes(s))));renderList();renderMarkers();updateCounts();syncRiskButtons();}
function updateCounts(){const c=state.companies;$('countAll').textContent=fmt(c.length);$('countHigh').textContent=fmt(c.filter(x=>x.grade==='고위험').length);$('countMid').textContent=fmt(c.filter(x=>x.grade==='주의').length);$('countLow').textContent=fmt(c.filter(x=>x.grade==='일반').length);$('countMapped').textContent=fmt(c.filter(x=>x.lat!==null).length);$('rowsLabel').textContent=`${state.results.length}개 표시`; if($('mobileMapCount'))$('mobileMapCount').textContent=`표시 대상 ${state.results.length}개 · 지도 ${state.results.filter(x=>x.lat!==null).length}개`; }
function riskClass(grade){return grade==='고위험'?'high':grade==='주의'?'mid':grade==='일반'?'low':'gray'}
function renderList(){const list=$('companyList');list.replaceChildren();for(const c of state.results){const el=document.createElement('div');el.className='company'+(c.id===state.selected?' active':'');el.innerHTML=`<i class="risk-dot ${riskClass(c.grade)}"></i><div><div class="company-name">${esc(c.name)}</div><div class="company-sub">${esc(c.id)} · ${esc(c.district)} · ${esc(c.grade)} ${c.score}점</div><small>${c.lat===null?'⚪ 좌표 미확인':'📍 지도 표시 가능'} · ${esc(c.status)}</small></div>`;el.onclick=()=>selectCompany(c.id,true);list.append(el);}}
function field(k,v){return `<div class="field"><label>${esc(k)}</label><strong>${esc(label(v))}</strong></div>`}
function renderDetail(){const c=state.companies.find(x=>x.id===state.selected);const badge=$('riskBadge');badge.className='pill '+(c?riskClass(c.grade):'');badge.textContent=c?`${c.grade} · ${c.score}점`:'미선택';if(!c){$('detailContent').innerHTML='<p class="empty">왼쪽 업체를 선택하거나 지도 마커를 클릭하세요.</p>';return;}
 const byYear=c.perf.map(p=>`<div class="statline"><span>${esc(p['실적연도'])}년 · ${esc(p['집계범위'])}</span><b>${fmt(num(p['환전금액(만원)']))}만원</b></div>`).join('')||'<p>자료 없음</p>';
 const checks=c.checks.map(p=>`<div class="statline"><span>${esc(p['점검일'])} · ${esc(p['점검구분'])}</span><b>${esc(p['결과분류'])}</b></div><p class="muted">${esc(p['주요결과(가상)'])} / ${esc(p['후속조치'])}</p>`).join('')||'<p>자료 없음</p>';
 $('detailContent').innerHTML=`<h3 class="detail-title">${esc(c.name)}</h3><div class="detail-sub">${esc(c.id)} · 100% 가상 업체</div>${field('대표자',c.owner)}${field('소재지',c.address)}${field('영업상태',c.status)}${field('등록일',c.registered)}${field('좌표 검증',c.geoStatus)}<div class="section-title">위험도 평가</div><div class="score"><b>${c.score}</b><span>/ 100점 · ${esc(c.grade)}</span></div>${field('실적변동',norm(c.risk['실적변동점수(최대30)'])+' / 30')}${field('점검경과',norm(c.risk['점검경과점수(최대20)'])+' / 20')}${field('위반이력',norm(c.risk['위반이력점수(최대30)'])+' / 30')}${field('영업상태',norm(c.risk['휴폐업등점수(최대20)'])+' / 20')}<div class="section-title">환전실적</div>${byYear}<div class="section-title">점검이력</div>${checks}`;
}
function selectCompany(id,pan){state.selected=id; if(window.innerWidth<=700){if($('main').dataset.mobileTab==='map'){document.body.classList.add('sheet-open');}else{setMobileTab('detail');}}const c=state.companies.find(x=>x.id===id);renderDetail();renderList();if(c&&pan&&state.map&&c.lat!==null){state.map.panTo(new kakao.maps.LatLng(c.lat,c.lng));state.map.setLevel(Math.min(state.map.getLevel(),5));}}
function clearMarkers(){if(state.clusterer)state.clusterer.clear();for(const m of state.markers)m.setMap(null);state.markers=[]}
function renderMarkers(){
 if(!state.map)return;
 clearMarkers();
 const overlaps={};
 // 동일 주소는 작은 링 형태로 분리하되, 축소 지도에서는 클러스터 숫자로 표시합니다.
 const ordered=state.results.filter(c=>c.lat!==null).sort((a,b)=>({일반:0,주의:1,고위험:2}[a.grade]||0)-({일반:0,주의:1,고위험:2}[b.grade]||0));
 for(const c of ordered){
  const pair=c.lat.toFixed(6)+','+c.lng.toFixed(6);
  const offset=overlaps[pair]||0;overlaps[pair]=offset+1;
  const ring=offset===0?0:Math.ceil(offset/8);
  const angle=(offset%8)*Math.PI/4;
  const deltaLat=ring*0.00016*Math.sin(angle),deltaLng=ring*0.00020*Math.cos(angle);
  const colors={high:'#d94658',mid:'#e6a329',low:'#21ad80',gray:'#8493a2'};
  const color=colors[riskClass(c.grade)];
  const svg=`<svg xmlns="http://www.w3.org/2000/svg" width="30" height="39" viewBox="0 0 30 39"><path d="M15 1C7 1 1 7 1 15c0 10 14 23 14 23s14-13 14-23C29 7 23 1 15 1Z" fill="${color}" stroke="white" stroke-width="2"/><circle cx="15" cy="15" r="6" fill="white"/></svg>`;
  const image=new kakao.maps.MarkerImage('data:image/svg+xml;charset=utf-8,'+encodeURIComponent(svg),new kakao.maps.Size(30,39),{offset:new kakao.maps.Point(15,39)});
  const marker=new kakao.maps.Marker({position:new kakao.maps.LatLng(c.lat+deltaLat,c.lng+deltaLng),image,title:c.name,zIndex:c.grade==='고위험'?30:c.grade==='주의'?20:10});
  kakao.maps.event.addListener(marker,'click',()=>selectCompany(c.id,true));
  state.markers.push(marker);
 }
 if(state.clusterer)state.clusterer.addMarkers(state.markers);
 else for(const m of state.markers)m.setMap(state.map);
}
function fitAllMarkers(){
 if(!state.map)return;
 const found=state.results.filter(c=>c.lat!==null);
 if(!found.length){state.map.setCenter(new kakao.maps.LatLng(37.3219,126.8309));state.map.setLevel(9);return;}
 const bounds=new kakao.maps.LatLngBounds();
 for(const c of found)bounds.extend(new kakao.maps.LatLng(c.lat,c.lng));
 state.map.setBounds(bounds,42,42,42,42);
 document.body.classList.remove('sheet-open');
}
function syncRiskButtons(){
 const current=$('risk').value;
 for(const btn of document.querySelectorAll('[data-grade]')){
  btn.classList.toggle('active',btn.dataset.grade===current);
  btn.setAttribute('aria-pressed',String(btn.dataset.grade===current));
 }
 if($('mobileRisk'))$('mobileRisk').value=current;
}
function loadKakao(key){if(state.map){setStatus('이미 지도가 연결되었습니다. 키 변경은 페이지 새로고침 후 적용하세요.');return;}$('keyStatus').textContent='지도 SDK 불러오는 중…';const script=document.createElement('script');script.src=`https://dapi.kakao.com/v2/maps/sdk.js?appkey=${encodeURIComponent(key)}&libraries=services,clusterer&autoload=false`;script.onload=()=>{if(!window.kakao?.maps){$('keyStatus').textContent='Kakao SDK 로드 실패. 도메인과 키를 확인하세요.';return;}kakao.maps.load(()=>{try{state.map=new kakao.maps.Map($('map'),{center:new kakao.maps.LatLng(37.3219,126.8309),level:8});state.geocoder=new kakao.maps.services.Geocoder();state.clusterer=new kakao.maps.MarkerClusterer({map:state.map,averageCenter:true,minLevel:5,disableClickZoom:false,styles:[{width:'42px',height:'42px',background:'#195d93eb',border:'3px solid white',borderRadius:'50%',color:'#fff',textAlign:'center',fontWeight:'bold',lineHeight:'36px',boxShadow:'0 2px 10px #1234'}]});$('mapPlaceholder').style.display='none';$('modal').classList.remove('open');renderMarkers();setStatus('카카오맵 연결 완료 · 가상업체 주소 변환 준비');if(state.companies.length)geocodeAll();}catch(e){$('keyStatus').textContent='지도 초기화 실패: '+e.message;}})};script.onerror=()=>{$('keyStatus').textContent='SDK 로드 실패: 키·도메인·기관망 접근 허용 여부를 확인하세요.'};document.head.appendChild(script);}
const sleep=ms=>new Promise(r=>setTimeout(r,ms));
function geocodeOne(address){return new Promise(resolve=>{let settled=false;const timeout=setTimeout(()=>{if(!settled){settled=true;resolve(null)}},8500);try{state.geocoder.addressSearch(address,(rows,status)=>{if(settled)return;settled=true;clearTimeout(timeout);if(status!==kakao.maps.services.Status.OK||!rows?.length)return resolve(null);const r=rows[0],lat=Number(r.y),lng=Number(r.x);if(!(lat>=37.20&&lat<=37.45&&lng>=126.55&&lng<=127.05))return resolve(null);resolve({lat,lng});});}catch(e){if(!settled){settled=true;clearTimeout(timeout);resolve(null)}}});}
async function geocodeAll(){
 if(state.busy)return;
 if(!state.geocoder)return setStatus('먼저 카카오맵 API 키로 지도를 연결하세요.');
 if(!state.companies.length)return setStatus('가상 데이터가 준비되지 않았습니다.');
 state.busy=true;$('geocodeBtn').disabled=true;
 const pending=state.companies.filter(x=>x.lat===null);
 const groups=new Map();for(const c of pending){if(!groups.has(c.address))groups.set(c.address,[]);groups.get(c.address).push(c);}
 let ok=0,fail=0,i=0;
 for(const [address,group] of groups){
  i++;setStatus(`안산시 전역 위치 확인 중 ${i}/${groups.size}개 주소 · 성공 ${ok} / 실패 ${fail}`);
  const pos=await geocodeOne(address);
  for(const c of group){if(pos){Object.assign(c,pos,{geoStatus:'확인완료'});state.coords[c.id]={address:c.address,...pos};ok++;}else{c.geoStatus='변환실패';fail++;}}
  filter();await sleep(120);
 }
 try{localStorage.setItem('fxgis_geocache_v1',JSON.stringify(state.coords))}catch(e){}
 state.busy=false;$('geocodeBtn').disabled=false;filter();
 setStatus(`가상업체 ${state.companies.length}개 · 지도표시 ${state.companies.filter(c=>c.lat!==null).length}개\n주소기준점은 시험용이며 실제 업체 위치가 아닙니다.`);
}
function csvEscape(v){return '"'+norm(v).replace(/"/g,'""')+'"'}
function exportCsv(){if(!state.companies.length)return setStatus('먼저 엑셀을 불러오세요.');const heads=['업체ID','상호','대표자','주소','구','영업상태','위험등급','위험점수','위도','경도','좌표상태'];const lines=[heads.map(csvEscape).join(',')];for(const c of state.companies)lines.push([c.id,c.name,c.owner,c.address,c.district,c.status,c.grade,c.score,c.lat??'',c.lng??'',c.geoStatus].map(csvEscape).join(','));const blob=new Blob(['\uFEFF'+lines.join('\r\n')],{type:'text/csv;charset=utf-8'});const url=URL.createObjectURL(blob);const a=document.createElement('a');a.href=url;a.download='안산_환전업자_좌표변환결과.csv';a.click();setTimeout(()=>URL.revokeObjectURL(url),1000);}
$('fileInput').addEventListener('change',async e=>{const file=e.target.files?.[0];if(!file)return;if(!window.XLSX)return setStatus('엑셀 파서 로드 실패. SheetJS CDN 접근을 확인하세요.');try{setStatus('엑셀 파싱 중…');const bytes=await file.arrayBuffer();const wb=XLSX.read(bytes,{type:'array',cellDates:false});ingest(wb);}catch(err){setStatus('엑셀 읽기 실패: '+err.message);}e.target.value='';});
for(const id of ['search','district','risk'])$(id).addEventListener(id==='search'?'input':'change',filter);
$('geocodeBtn').onclick=geocodeAll;$('exportBtn').onclick=exportCsv;
function showModal(){$('apiKey').value=state.key;$('modal').classList.add('open');$('keyStatus').textContent='';}
$('configBtn').onclick=showModal;$('setupHere').onclick=showModal;$('closeModal').onclick=()=>$('modal').classList.remove('open');
$('connectBtn').onclick=()=>{const key=$('apiKey').value.trim();if(!key){$('keyStatus').textContent='JavaScript 키를 입력하세요.';return;}state.key=key;try{localStorage.setItem('fxgis_kakao_js_key',key)}catch(e){}loadKakao(key)};
$('forgetKey').onclick=()=>{state.key='';$('apiKey').value='';localStorage.removeItem('fxgis_kakao_js_key');$('keyStatus').textContent='저장된 키를 삭제했습니다. 이미 연결된 지도는 새로고침 후 해제됩니다.'};
try{state.coords=JSON.parse(localStorage.getItem('fxgis_geocache_v1')||'{}')||{}}catch(e){state.coords={}}
updateCounts();if(state.key)loadKakao(state.key);

function setMobileTab(tab){
 $('main').dataset.mobileTab=tab;
 document.body.classList.remove('sheet-open');
 for(const [id,t] of [['mobileMapBtn','map'],['mobileListBtn','list'],['mobileDetailBtn','detail']])$(id).classList.toggle('active',t===tab);
 for(const [id,t] of [['bottomMap','map'],['bottomList','list'],['bottomDetail','detail']])$(id).classList.toggle('active',t===tab);
 if(tab==='map'&&state.map)setTimeout(()=>state.map.relayout(),70);
}
$('mobileMapBtn').onclick=()=>setMobileTab('map');
$('mobileListBtn').onclick=()=>setMobileTab('list');
$('mobileDetailBtn').onclick=()=>setMobileTab('detail');
$('bottomMap').onclick=()=>setMobileTab('map');
$('bottomList').onclick=()=>setMobileTab('list');
$('bottomDetail').onclick=()=>setMobileTab('detail');
$('mobileSheetClose').onclick=()=>document.body.classList.remove('sheet-open');
$('mobileAllBtn').onclick=fitAllMarkers;
$('mobileSearch').addEventListener('input',e=>{$('search').value=e.target.value;filter();});
$('mobileDistrict').addEventListener('change',e=>{$('district').value=e.target.value;filter();});
$('mobileRisk').addEventListener('change',e=>{$('risk').value=e.target.value;filter();});
$('mobileFilterBtn').onclick=()=>{const e=$('mobileFilters');e.hidden=!e.hidden;};
$('search').addEventListener('input',()=>{$('mobileSearch').value=$('search').value;});
$('district').addEventListener('change',()=>{$('mobileDistrict').value=$('district').value;});
$('risk').addEventListener('change',()=>{$('mobileRisk').value=$('risk').value;});
window.addEventListener('keydown',e=>{if(e.key==='Escape')document.body.classList.remove('sheet-open');});

(async function bootDemo(){
 try{
  const resp=await fetch('./demo-data.json?v=0.4.1');if(!resp.ok)throw Error('HTTP '+resp.status);
  const data=await resp.json();const wb={Sheets:{}};
  for(const [name,rows] of Object.entries(data))wb.Sheets[name]={__rows:rows};
  // readTable() supports compact pre-parsed demo data and ordinary XLSX imports.
  ingest(wb);
 }catch(e){setStatus('기본 가상 데이터 로드 오류: '+e.message+' · 엑셀 업로드는 가능합니다.');}
})();

for(const btn of document.querySelectorAll('[data-grade]'))btn.addEventListener('click',()=>{const grade=btn.dataset.grade;$('risk').value=grade;if($('mobileRisk'))$('mobileRisk').value=grade;filter();if(grade==='좌표 미확인')setMobileTab('list');else {setMobileTab('map');fitAllMarkers();}});
