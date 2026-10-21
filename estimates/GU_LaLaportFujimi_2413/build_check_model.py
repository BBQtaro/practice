# -*- coding: utf-8 -*-
"""根拠確認用モデル（HTML）を作る: 3D モデル・平面図・数量比較を 1 ファイルにまとめる

python3 build_check_model.py  ->  GU富士見_2413_根拠確認モデル.html
画像を使わず takeoff_geometry.py の座標だけで描くので、単体で開ける。
"""
import json
import takeoff_geometry as g

OUT = "GU富士見_2413_根拠確認モデル.html"

# 見積 v2（2021 実績数量ベース）と図面拾いの比較
COMPARE = [
    ("LGS ST100 壁下地（㎡）", 691, g.summary()["LGS ST100（㎡）"], "実績は壁面什器背面・FR・レジバックの下地を含むとみられる。図面拾いは区画外周＋売場外周のみ"),
    ("LGS ST65 間仕切（㎡）", 285, g.summary()["LGS ST65 両面間仕切（㎡）"], "図面拾いはSW背面・FRブース間仕切を含む"),
    ("柱 LGS＋PB（㎡）", 156, g.summary()["柱 LGS＋PB（㎡）"], "図面拾いは売場独立柱7本のみ（壁付柱を含めると増）"),
    ("PB12.5（壁・柱, ㎡）", 965 + 156, g.summary()["PB12.5（壁・柱・下がり壁, ㎡）"], "ほぼ一致"),
    ("天井 AEP（㎡）", 1065, g.summary()["天井 AEP（全室, ㎡）"], "一致"),
    ("壁 AEP N-87（壁・柱・ファサード, ㎡）", 645 + 64 + 59, g.finish_totals()["AEP N-87"], "図面拾いは什器で隠れる面も全面計上。実績は塗装範囲を絞っている可能性"),
    ("白クロス（㎡）", 131, g.finish_totals()["白クロス LW-4630/4635"], "FRブース間仕切の両面分で図面拾いが多い"),
    ("黒クロス（㎡）", 55, g.finish_totals()["黒クロス LW-4207"], "実績は標準化項目「ベース壁面クロス55㎡」"),
    ("化粧PB 木目（㎡）", 273, g.finish_totals()["化粧PB（木目・桐）"], "図面拾いは間仕切のみ。休憩室・店長室の外周壁内面も化粧PBなら増 → 要確認"),
    ("床 ES3001 施工（㎡）", 876, g.summary()["床 ES3001（売場＋レジ＋FRブース外, ㎡）"], "ほぼ一致"),
    ("床 P-41 施工（㎡）", 169, g.summary()["床 P-41（BOH＋SW, ㎡）"], "図面ではSW 21.6㎡もP-41 → 見積に追加検討"),
    ("柱巻きミラー（枚）", 14, len(g.COLUMNS) * 2, "一致（独立柱7本×2面）"),
    ("ダウンライト 取付（ヶ所）", 396, 396, "照明器具表の総数量"),
    ("コンセント 取付（ヶ所）", 87, 87, "照明器具表の総数量"),
]

data = {
    "zones": g.zone_rows(),
    "excluded": [dict(id=i, name=n, pts=p) for i, n, p in g.EXCLUDED],
    "walls": [dict(r, faces_detail=g.FACE_FINISH.get(r["id"], []), segment=r["id"] in g.SEGMENT_WALLS) for r in g.wall_rows()],
    "columns": [dict(name=f"{x}×{y}", x=g.GRID_X[x], y=g.GRID_Y[y], size=g.COLUMN_SIZE) for x, y in g.COLUMNS],
    "grid_x": g.GRID_X, "grid_y": g.GRID_Y,
    "summary": g.summary(), "finish": g.finish_totals(),
    "fixtures": [dict(name=n, qty=q, src=s) for n, q, s in g.FIXTURES],
    "compare": [dict(item=a, est=b, drw=c, note=d) for a, b, c, d in COMPARE],
}

HTML = r"""<!doctype html>
<html lang="ja">
<head>
<meta charset="utf-8">
<meta name="viewport" content="width=device-width, initial-scale=1">
<title>GU富士見 根拠確認モデル</title>
<style>
:root{--bg:#f6f7f9;--panel:#ffffff;--ink:#1f2933;--sub:#5f6b7a;--line:#d9dee5;--accent:#c53030;--hi:#fff4c2}
@media (prefers-color-scheme: dark){:root:not([data-theme="light"]){--bg:#15181d;--panel:#1e232a;--ink:#e6e9ee;--sub:#9aa5b1;--line:#2f3640;--accent:#fc8181;--hi:#3d3a22}}
:root[data-theme="dark"]{--bg:#15181d;--panel:#1e232a;--ink:#e6e9ee;--sub:#9aa5b1;--line:#2f3640;--accent:#fc8181;--hi:#3d3a22}
*{box-sizing:border-box}
body{margin:0;background:var(--bg);color:var(--ink);font-family:"Hiragino Sans","Noto Sans JP","Yu Gothic",sans-serif;font-size:14px}
header{padding:14px 16px;border-bottom:1px solid var(--line);background:var(--panel)}
header h1{margin:0;font-size:18px}
header p{margin:4px 0 0;color:var(--sub);font-size:12px}
nav{display:flex;gap:4px;padding:8px 16px;background:var(--panel);border-bottom:1px solid var(--line);flex-wrap:wrap}
nav button{border:1px solid var(--line);background:transparent;color:var(--ink);padding:6px 12px;border-radius:6px;cursor:pointer;font-size:13px}
nav button.on{background:var(--ink);color:var(--panel)}
main{display:grid;grid-template-columns:minmax(0,1fr) 340px;gap:12px;padding:12px 16px}
@media (max-width:900px){main{grid-template-columns:1fr}}
.card{background:var(--panel);border:1px solid var(--line);border-radius:8px;padding:12px;min-width:0}
#view3d{height:62vh;min-height:360px;border-radius:6px;overflow:hidden;background:#dfe4ea}
#svgwrap{overflow:auto}
svg text{font-family:inherit}
.toggles label{display:flex;align-items:center;gap:6px;margin:3px 0;font-size:13px}
.sw{display:inline-block;width:14px;height:14px;border-radius:3px;border:1px solid #0003}
#info{font-size:13px;line-height:1.6}
#info .t{font-weight:700;font-size:15px;margin-bottom:4px}
table{border-collapse:collapse;width:100%;font-size:12.5px}
th,td{border-bottom:1px solid var(--line);padding:5px 6px;text-align:left;vertical-align:top}
th{color:var(--sub);font-weight:600}
td.n{text-align:right;font-variant-numeric:tabular-nums;white-space:nowrap}
tr.sel{background:var(--hi)}
.big{color:var(--accent);font-weight:700}
.tab{display:none}.tab.on{display:block}
.scroll{overflow-x:auto}
small{color:var(--sub)}
</style>
</head>
<body>
<header>
<h1>GU ららぽーと富士見店（2413区画）根拠確認モデル</h1>
<p>施工図一式（A03 平面図・A04 床伏図・A05 天井伏図・A16〜A21 展開図・照明器具表）を通り芯基準で読み取った拾い出し。床・壁・柱をクリックすると数量と根拠が右に出ます。</p>
</header>
<nav>
<button data-tab="t3d" class="on">3Dモデル</button>
<button data-tab="tplan">平面（拾い範囲）</button>
<button data-tab="tqty">数量・見積との比較</button>
</nav>
<main>
<section class="card">
  <div id="t3d" class="tab on"><div id="view3d"></div><small>ドラッグ：回転／ホイール：拡大縮小／右ドラッグ：移動。店頭（Y6側）が手前。</small></div>
  <div id="tplan" class="tab"><div id="svgwrap"></div></div>
  <div id="tqty" class="tab">
    <h3>見積 v2 と図面拾いの比較</h3>
    <div class="scroll"><table id="cmp"></table></div>
    <h3>床（室別）</h3><div class="scroll"><table id="zt"></table></div>
    <h3>壁（ID別）</h3><div class="scroll"><table id="wt"></table></div>
    <h3>壁仕上げ（面別合計）</h3><div class="scroll"><table id="ft"></table></div>
    <h3>器具・コンセント</h3><div class="scroll"><table id="xt"></table></div>
  </div>
</section>
<aside class="card">
  <div class="toggles" id="toggles"></div>
  <hr style="border:none;border-top:1px solid var(--line)">
  <div id="info"><div class="t">要素を選択してください</div><small>3D・平面・表のどれからでも選べます。</small></div>
</aside>
</main>
<script>const D = __DATA__;</script>
<script src="https://cdnjs.cloudflare.com/ajax/libs/three.js/r128/three.min.js"></script>
<script src="https://cdn.jsdelivr.net/npm/three@0.128.0/examples/js/controls/OrbitControls.js"></script>
<script>
const fmt=n=>typeof n==="number"?n.toLocaleString("ja-JP",{maximumFractionDigits:1}):n;
const show={zone:true,wall:true,column:true,excluded:true};
let selected=null;
// ---- 情報パネル
function info(o){
  selected=o.id;
  const el=document.getElementById("info");
  if(o.kind==="zone"){
    el.innerHTML=`<div class="t">${o.id} ${o.name}</div>床仕上：${o.finish}<br>正式面積（平面図面積表）：<span class="big">${fmt(o.official)} ㎡</span><br>図上概算：${fmt(o.area)} ㎡（差 ${fmt(o.official?((o.area/o.official-1)*100):0)}%）<br><small>根拠：A04 床伏図・A03 面積表／積算根拠図 p.2</small>`;
  }else if(o.kind==="wall"){
    const f=o.faces_detail.map(([k,n])=>`${k} ×${n}面`).join("<br>");
    el.innerHTML=`<div class="t">${o.id} ${o.name}</div>下地：${o.base}<br>長さ：<span class="big">${fmt(o.length)} m</span>　高さ：${fmt(o.height)} mm<br>面積（片面）：<span class="big">${fmt(o.area)} ㎡</span><br>仕上：<br>${f}<br><small>根拠：A03 平面図（長さ）・${o.ref}（仕上・高さ）／積算根拠図 p.3</small>`;
  }else if(o.kind==="column"){
    el.innerHTML=`<div class="t">柱 ${o.name}</div>柱巻きミラー 2面（W600×H1800）＋LGS・PB・AEP<br><small>根拠：A03 平面図・展開2/14／売場独立柱 7本</small>`;
  }else{el.innerHTML=`<div class="t">${o.name}</div>A工事範囲のため計上しない`;}
  document.querySelectorAll("tr[data-id]").forEach(r=>r.classList.toggle("sel",r.dataset.id===o.id));
  highlight3d(o.id);
}
// ---- トグル
const tg=document.getElementById("toggles");
[["zone","床ゾーン"],["wall","壁"],["column","柱"],["excluded","対象外（階段室）"]].forEach(([k,l])=>{
  const lab=document.createElement("label");
  lab.innerHTML=`<input type="checkbox" checked> ${l}`;
  lab.querySelector("input").onchange=e=>{show[k]=e.target.checked;apply3d();drawSVG();};
  tg.appendChild(lab);
});
// ---- タブ
document.querySelectorAll("nav button").forEach(b=>b.onclick=()=>{
  document.querySelectorAll("nav button").forEach(x=>x.classList.toggle("on",x===b));
  document.querySelectorAll(".tab").forEach(t=>t.classList.toggle("on",t.id===b.dataset.tab));
  if(b.dataset.tab==="t3d")resize3d();
});
// ---- 表
function table(id,head,rows){
  const t=document.getElementById(id);
  t.innerHTML="<tr>"+head.map(h=>`<th>${h}</th>`).join("")+"</tr>"+rows.map(r=>`<tr ${r.id?`data-id="${r.id}"`:""}>${r.cells.map((c,i)=>`<td class="${typeof c==="number"?"n":""}">${fmt(c)}</td>`).join("")}</tr>`).join("");
}
table("cmp",["項目","見積v2（実績数量）","図面拾い","差","メモ"],D.compare.map(c=>({cells:[c.item,c.est,c.drw,(c.est?((c.drw/c.est-1)*100).toFixed(0)+"%":"-"),c.note]})));
table("zt",["ID","室名","床仕上","正式㎡","図上㎡"],D.zones.map(z=>({id:z.id,cells:[z.id,z.name,z.finish,z.official,z.area]})));
table("wt",["ID","名称","下地","長さm","高さmm","片面㎡","参照"],D.walls.map(w=>({id:w.id,cells:[w.id,w.name,w.base,w.length,w.height,w.area,w.ref]})));
table("ft",["仕上","面積㎡"],Object.entries(D.finish).map(([k,v])=>({cells:[k,v]})));
table("xt",["品目","数量","根拠"],D.fixtures.map(f=>({cells:[f.name,f.qty,f.src]})));
document.querySelectorAll("tr[data-id]").forEach(r=>r.onclick=()=>{
  const id=r.dataset.id;const o=D.zones.find(z=>z.id===id)?{...D.zones.find(z=>z.id===id),kind:"zone"}:{...D.walls.find(w=>w.id===id),kind:"wall"};info(o);
});
// ---- 平面 SVG
function drawSVG(){
  const minX=-2000,minY=-5000,W=38000,H=48000,sc=0.02;
  let s=`<svg viewBox="${minX} ${minY} ${W} ${H}" width="${W*sc}" style="max-width:100%;height:auto;background:#fff">`;
  for(const[k,v]of Object.entries(D.grid_x))s+=`<line x1="${v}" y1="${minY}" x2="${v}" y2="${minY+H}" stroke="#bbb" stroke-dasharray="400 200" stroke-width="30"/><text x="${v}" y="${minY+900}" font-size="700" text-anchor="middle" fill="#666">${k}</text>`;
  for(const[k,v]of Object.entries(D.grid_y))s+=`<line x1="${minX}" y1="${v}" x2="${minX+W}" y2="${v}" stroke="#bbb" stroke-dasharray="400 200" stroke-width="30"/><text x="${minX+300}" y="${v-150}" font-size="700" fill="#666">${k}</text>`;
  if(show.zone)D.zones.forEach(z=>{s+=`<polygon data-k="zone" data-id="${z.id}" points="${z.pts.map(p=>p.join(",")).join(" ")}" fill="${z.color}" fill-opacity="${selected===z.id?0.75:0.4}" stroke="${z.color}" stroke-width="60" style="cursor:pointer"/>`;
    const LP={F1:[17500,22000],F2:[30000,9000],F8:[31200,22000]};const cx=LP[z.id]?LP[z.id][0]:z.pts.reduce((a,p)=>a+p[0],0)/z.pts.length,cy=LP[z.id]?LP[z.id][1]:z.pts.reduce((a,p)=>a+p[1],0)/z.pts.length;
    s+=`<text x="${cx}" y="${cy}" font-size="650" text-anchor="middle" fill="#222" pointer-events="none">${z.id} ${z.name} ${z.official}㎡</text>`;});
  if(show.excluded)D.excluded.forEach(e=>s+=`<polygon points="${e.pts.map(p=>p.join(",")).join(" ")}" fill="#777" fill-opacity=".45"/>`);
  if(show.wall)D.walls.forEach(w=>{
    const segs=w.segment?w.pts.reduce((a,p,i)=>i%2?a:[...a,[p,w.pts[i+1]]],[]):[w.pts];
    segs.forEach(seg=>s+=`<polyline data-k="wall" data-id="${w.id}" points="${seg.map(p=>p.join(",")).join(" ")}" fill="none" stroke="${w.color}" stroke-width="${selected===w.id?420:250}" stroke-linecap="square" style="cursor:pointer"/>`);
    const p=w.pts[Math.floor(w.pts.length/2)];s+=`<text x="${p[0]+250}" y="${p[1]+250}" font-size="560" fill="#c53030" pointer-events="none">${w.id}</text>`;});
  if(show.column)D.columns.forEach(c=>s+=`<rect data-k="column" data-id="${c.name}" x="${c.x-c.size/2}" y="${c.y-c.size/2}" width="${c.size}" height="${c.size}" fill="#38a169" style="cursor:pointer"/>`);
  s+=`<text x="${minX+500}" y="${minY+H-500}" font-size="800" fill="#444">店頭（Y6側）↓　1/200 図面を通り芯基準で読み取り</text></svg>`;
  const wrap=document.getElementById("svgwrap");wrap.innerHTML=s;
  wrap.querySelectorAll("[data-k]").forEach(el=>el.onclick=()=>{const k=el.dataset.k,id=el.dataset.id;
    if(k==="zone")info({...D.zones.find(z=>z.id===id),kind:"zone"});
    else if(k==="wall")info({...D.walls.find(w=>w.id===id),kind:"wall"});
    else info({kind:"column",name:id,id});drawSVG();});
}
drawSVG();
// ---- 3D
let scene,cam,ren,ctl,meshes=[];
function init3d(){
  if(typeof THREE==="undefined"){document.getElementById("view3d").innerHTML='<p style="padding:16px">3D表示用のライブラリを読み込めませんでした（オフライン？）。平面タブと数量タブは使えます。</p>';return;}
  const el=document.getElementById("view3d");
  scene=new THREE.Scene();scene.background=new THREE.Color(0xdfe4ea);
  cam=new THREE.PerspectiveCamera(45,el.clientWidth/el.clientHeight,10,500000);
  cam.position.set(17000,-80000,38000);
  ren=new THREE.WebGLRenderer({antialias:true});ren.setPixelRatio(window.devicePixelRatio);el.appendChild(ren.domElement);
  ctl=new THREE.OrbitControls(cam,ren.domElement);ctl.target.set(17000,-19000,0);
  scene.add(new THREE.AmbientLight(0xffffff,.65));const dl=new THREE.DirectionalLight(0xffffff,.6);dl.position.set(20000,-60000,80000);scene.add(dl);
  // 座標: x=右, y=-(図面y)（店頭が手前）, z=高さ
  const T=(p)=>new THREE.Vector2(p[0],-p[1]);
  D.zones.forEach(z=>{const sh=new THREE.Shape(z.pts.map(T));const m=new THREE.Mesh(new THREE.ShapeGeometry(sh),new THREE.MeshLambertMaterial({color:z.color,side:THREE.DoubleSide,transparent:true,opacity:.85}));m.userData={kind:"zone",...z};meshes.push(m);scene.add(m);});
  D.excluded.forEach(e=>{const m=new THREE.Mesh(new THREE.ShapeGeometry(new THREE.Shape(e.pts.map(T))),new THREE.MeshLambertMaterial({color:0x777777}));m.position.z=5;m.userData={kind:"excluded",...e};meshes.push(m);scene.add(m);});
  D.walls.forEach(w=>{
    const segs=w.segment?w.pts.reduce((a,p,i)=>i%2?a:[...a,[p,w.pts[i+1]]],[]):w.pts.slice(1).map((p,i)=>[w.pts[i],p]);
    const z0=w.id==="W17"?3750-650:0;
    segs.forEach(([a,b])=>{const dx=b[0]-a[0],dy=-(b[1]-a[1]),L=Math.hypot(dx,dy);
      const m=new THREE.Mesh(new THREE.BoxGeometry(L,150,w.height),new THREE.MeshLambertMaterial({color:w.color,transparent:true,opacity:.92}));
      m.position.set((a[0]+b[0])/2,-(a[1]+b[1])/2,z0+w.height/2);m.rotation.z=Math.atan2(dy,dx);m.userData={kind:"wall",...w};meshes.push(m);scene.add(m);});
  });
  D.columns.forEach(c=>{const m=new THREE.Mesh(new THREE.BoxGeometry(c.size,c.size,3750),new THREE.MeshLambertMaterial({color:0x38a169}));m.position.set(c.x,-c.y,1875);m.userData={kind:"column",name:c.name,id:c.name};meshes.push(m);scene.add(m);});
  const ray=new THREE.Raycaster(),mv=new THREE.Vector2();
  ren.domElement.addEventListener("click",e=>{const r=ren.domElement.getBoundingClientRect();mv.x=((e.clientX-r.left)/r.width)*2-1;mv.y=-((e.clientY-r.top)/r.height)*2+1;ray.setFromCamera(mv,cam);
    const hit=ray.intersectObjects(meshes.filter(m=>m.visible))[0];if(hit)info(hit.object.userData);});
  resize3d();apply3d();(function loop(){requestAnimationFrame(loop);ctl.update();ren.render(scene,cam);})();
}
function resize3d(){if(!ren)return;const el=document.getElementById("view3d");cam.aspect=el.clientWidth/el.clientHeight;cam.updateProjectionMatrix();ren.setSize(el.clientWidth,el.clientHeight);}
function apply3d(){meshes.forEach(m=>m.visible=!!show[m.userData.kind]);}
function highlight3d(id){meshes.forEach(m=>{if(m.material&&m.material.emissive)m.material.emissive.setHex((m.userData.id===id)?0x553300:0x000000);});}
window.addEventListener("resize",resize3d);
window.addEventListener("load",init3d);
</script>
</body>
</html>
"""

with open(OUT, "w", encoding="utf-8") as f:
    f.write(HTML.replace("__DATA__", json.dumps(data, ensure_ascii=False)))
print("saved", OUT)
