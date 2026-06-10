import { useState, useEffect, useRef } from "react";
import * as XLSX from "xlsx";
import { BarChart, Bar, LineChart, Line, XAxis, YAxis, CartesianGrid, Tooltip, Legend, ResponsiveContainer } from "recharts";

// ─────────────────────────────────────────────
// CONSTANTES
// ─────────────────────────────────────────────
const LOCALES = [
  { id: "local1", nombre: "Cornella",  emoji: "🏪", color: "#c9a84c" },
  { id: "local2", nombre: "Badalona",  emoji: "🏬", color: "#6a9fd8" },
];

// Tipos de movimiento con su efecto en el cálculo
const TIPOS_MOV = {
  venta:    { label: "Ventas efectivo", emoji: "💵", signo: +1, grupo: "ingreso", color: "#4caf82" },
  deposito: { label: "Depósito",        emoji: "🏦", signo: +1, grupo: "ingreso", color: "#6ac8d8" },
  gasto:    { label: "Gasto",           emoji: "🧾", signo: -1, grupo: "egreso",  color: "#c0503a" },
  retiro:   { label: "Retiro",          emoji: "💸", signo: -1, grupo: "egreso",  color: "#d48a3a" },
};

const CATS = {
  venta:    ["Ventas mostrador","Ventas delivery","Otros ingresos"],
  deposito: ["Depósito bancario","Cobro cliente","Transferencia recibida"],
  gasto:    ["Proveedores","Salarios","Alquiler","Servicios","Transporte","Otros gastos"],
  retiro:   ["Retiro propietario","Pago anticipado","Otros retiros"],
};

const MESES = ["Enero","Febrero","Marzo","Abril","Mayo","Junio","Julio","Agosto","Septiembre","Octubre","Noviembre","Diciembre"];
const DIAS_CORTOS = ["Dom","Lun","Mar","Mié","Jue","Vie","Sáb"];


// ─────────────────────────────────────────────
// HELPERS
// ─────────────────────────────────────────────
function getDateKey(date) { const d = date||new Date(); return d.toISOString().split("T")[0]; }

function getWeekNumber(dateStr) {
  const [y,m,d] = dateStr.split("-").map(Number);
  const date = new Date(y,m-1,d);
  const soy = new Date(y,0,1);
  const doy = Math.floor((date-soy)/86400000);
  return Math.ceil((doy+(soy.getDay()||7))/7);
}

function formatCurrency(val) {
  return new Intl.NumberFormat("es-ES",{style:"currency",currency:"EUR",maximumFractionDigits:2}).format(val||0);
}
function formatCurrencyShort(val) {
  if(Math.abs(val)>=1000000) return (val/1000000).toFixed(1)+"M";
  if(Math.abs(val)>=1000) return (val/1000).toFixed(0)+"K";
  return val;
}
function formatDate(dateStr) {
  const [y,m,d]=dateStr.split("-");
  return new Date(+y,+m-1,+d).toLocaleDateString("es-ES",{weekday:"long",year:"numeric",month:"long",day:"numeric"});
}
function formatDateShort(dateStr) {
  const [y,m,d]=dateStr.split("-"); const date=new Date(+y,+m-1,+d);
  return `${DIAS_CORTOS[date.getDay()]} ${d}/${m}`;
}

function skDay(localId,dateKey) { return `caja_${localId}_${dateKey}`; }

function emptyDay() {
  return { saldoInicial:0, cajaReal:null, movimientos:[], cerrado:false, nota:"" };
}

function loadDay(localId,dateKey) {
  try { const r=localStorage.getItem(skDay(localId,dateKey)); if(r) return JSON.parse(r); } catch{}
  return emptyDay();
}
function saveDay(localId,dateKey,data) {
  try { localStorage.setItem(skDay(localId,dateKey),JSON.stringify(data)); } catch{}
}
function getAllKeys(localId) {
  const prefix=`caja_${localId}_`, keys=[];
  for(let i=0;i<localStorage.length;i++){const k=localStorage.key(i); if(k&&k.startsWith(prefix)) keys.push(k.replace(prefix,""));}
  return keys.sort().reverse();
}
function getDaysInMonth(year,month) {
  const days=[],total=new Date(year,month,0).getDate();
  for(let d=1;d<=total;d++) days.push(`${year}-${String(month).padStart(2,"0")}-${String(d).padStart(2,"0")}`);
  return days;
}
function groupByWeek(days) {
  const w={};
  days.forEach(dk=>{ const n=getWeekNumber(dk); if(!w[n]) w[n]=[]; w[n].push(dk); });
  return w;
}

// Calcula todos los valores de un día
function calcDay(data) {
  const movs = data.movimientos||[];
  const ventas    = movs.filter(m=>m.tipo==="venta").reduce((s,m)=>s+m.monto,0);
  const depositos = movs.filter(m=>m.tipo==="deposito").reduce((s,m)=>s+m.monto,0);
  const gastos    = movs.filter(m=>m.tipo==="gasto").reduce((s,m)=>s+m.monto,0);
  const retiros   = movs.filter(m=>m.tipo==="retiro").reduce((s,m)=>s+m.monto,0);
  const cajaTeor  = (data.saldoInicial||0) + ventas + depositos - gastos - retiros;
  const cajaReal  = data.cajaReal ?? null;
  const diferencia= cajaReal !== null ? cajaReal - cajaTeor : null;
  return { ventas, depositos, gastos, retiros, cajaTeor, cajaReal, diferencia };
}

// ─────────────────────────────────────────────
// EXPORT EXCEL
// ─────────────────────────────────────────────
function exportarExcel() {
  const wb = XLSX.utils.book_new();
  LOCALES.forEach(local=>{
    const nombre = localStorage.getItem(`nombre_${local.id}`)||local.nombre;
    const keys = getAllKeys(local.id).sort();
    const rows = [];
    keys.forEach(dk=>{
      const d=loadDay(local.id,dk); const c=calcDay(d);
      rows.push({ Fecha:formatDate(dk), Semana:`S${getWeekNumber(dk)}`, Tipo:"─ RESUMEN ─",
        Ventas:c.ventas, Depósitos:c.depositos, Gastos:c.gastos, Retiros:c.retiros,
        "Caja Inicial":d.saldoInicial, "Caja Teórica":c.cajaTeor,
        "Caja Real":c.cajaReal??"–", Diferencia:c.diferencia??"–",
        Estado:d.cerrado?"Cerrada":"Abierta", Nota:d.nota||""});
      d.movimientos.forEach(m=>rows.push({
        Fecha:formatDate(dk),Semana:`S${getWeekNumber(dk)}`,
        Tipo:TIPOS_MOV[m.tipo]?.label||m.tipo,
        Ventas:m.tipo==="venta"?m.monto:"",Depósitos:m.tipo==="deposito"?m.monto:"",
        Gastos:m.tipo==="gasto"?m.monto:"",Retiros:m.tipo==="retiro"?m.monto:"",
        "Caja Inicial":"","Caja Teórica":"","Caja Real":"",Diferencia:"",
        Estado:"",Nota:m.descripcion}));
      rows.push({});
    });
    if(!rows.length) rows.push({Fecha:"Sin datos"});
    const ws=XLSX.utils.json_to_sheet(rows);
    ws["!cols"]=[{wch:32},{wch:8},{wch:18},{wch:12},{wch:12},{wch:12},{wch:12},{wch:12},{wch:12},{wch:12},{wch:10},{wch:10},{wch:28}];
    XLSX.utils.book_append_sheet(wb,ws,nombre.slice(0,31));
  });
  // Consolidado
  const allDates=[...new Set(getAllKeys("local1").concat(getAllKeys("local2")))].sort();
  const cRows=[];
  allDates.forEach(dk=>{
    LOCALES.forEach(local=>{
      const nombre=localStorage.getItem(`nombre_${local.id}`)||local.nombre;
      const d=loadDay(local.id,dk); const c=calcDay(d);
      cRows.push({Fecha:formatDate(dk),Semana:`S${getWeekNumber(dk)}`,Local:nombre,
        Ventas:c.ventas,Depósitos:c.depositos,Gastos:c.gastos,Retiros:c.retiros,
        "Caja Inicial":d.saldoInicial,"Caja Teórica":c.cajaTeor,
        "Caja Real":c.cajaReal??"–",Diferencia:c.diferencia??"–",Estado:d.cerrado?"Cerrada":"Abierta"});
    });
  });
  if(cRows.length){const ws=XLSX.utils.json_to_sheet(cRows);XLSX.utils.book_append_sheet(wb,ws,"Consolidado");}
  XLSX.writeFile(wb,`caja_${getDateKey()}.xlsx`);
}

// ─────────────────────────────────────────────
// COMPONENTES PEQUEÑOS
// ─────────────────────────────────────────────
function LocalNameEditor({local,onSave}) {
  const [editing,setEditing]=useState(false);
  const [val,setVal]=useState(()=>localStorage.getItem(`nombre_${local.id}`)||local.nombre);
  function save(){localStorage.setItem(`nombre_${local.id}`,val);onSave(val);setEditing(false);}
  if(editing) return(
    <span style={{display:"inline-flex",gap:4,alignItems:"center"}}>
      <input value={val} onChange={e=>setVal(e.target.value)} onKeyDown={e=>e.key==="Enter"&&save()}
        style={{background:"#0f0e0b",border:"1px solid #3a3520",borderRadius:6,padding:"3px 8px",color:"#e8e0cc",fontSize:13,width:120}} autoFocus/>
      <button onClick={save} style={{background:"transparent",border:"none",color:"#4caf82",cursor:"pointer",fontSize:14}}>✓</button>
    </span>);
  return <span style={{cursor:"pointer",borderBottom:"1px dashed #4a4335"}} onClick={()=>setEditing(true)} title="Clic para renombrar">{val}</span>;
}

function CustomTooltip({active,payload,label}) {
  if(!active||!payload?.length) return null;
  return(
    <div style={{background:"#1a1710",border:"1px solid #3a3520",borderRadius:8,padding:"10px 14px",fontSize:11}}>
      <div style={{color:"#8a7f5e",marginBottom:6}}>{label}</div>
      {payload.map(p=><div key={p.name} style={{color:p.color,marginBottom:2}}>{p.name}: {formatCurrency(p.value)}</div>)}
    </div>);
}

// ─────────────────────────────────────────────
// BLOQUE DESCUADRE
// ─────────────────────────────────────────────
function DescuadrePanel({calc,cajaReal,onSetCajaReal,editable}) {
  const [inputVal,setInputVal]=useState(cajaReal!==null?String(cajaReal):"");
  const [editando,setEditando]=useState(false);

  function guardar(){
    const v=parseFloat(inputVal);
    if(!isNaN(v)){onSetCajaReal(v);}
    setEditando(false);
  }

  const dif = calc.diferencia;
  const hayDif = dif!==null && dif!==0;
  const borderColor = dif===null?"#2e2b22":hayDif?"#c0503a33":"#2a5c3f33";
  const bgColor = dif===null?"#141210":hayDif?"#1f0e0e":"#0e1f16";

  return(
    <div style={{background:bgColor,border:`1px solid ${borderColor}`,borderRadius:10,padding:"12px 14px",marginBottom:14}}>
      <div style={{fontSize:10,letterSpacing:2,color:"#6a6047",textTransform:"uppercase",marginBottom:10}}>
        Arqueo de Caja
      </div>

      <div style={{display:"grid",gridTemplateColumns:"1fr 1fr 1fr",gap:8,marginBottom:10}}>
        <div style={{textAlign:"center"}}>
          <div style={{fontSize:9,color:"#5a5240",textTransform:"uppercase",letterSpacing:1,marginBottom:3}}>Caja Inicial</div>
          <div style={{fontSize:13,color:"#8a7f5e",fontWeight:"bold"}}>{formatCurrency(calc.cajaTeor - calc.ventas - calc.depositos + calc.gastos + calc.retiros)}</div>
        </div>
        <div style={{textAlign:"center"}}>
          <div style={{fontSize:9,color:"#5a5240",textTransform:"uppercase",letterSpacing:1,marginBottom:3}}>Caja Teórica</div>
          <div style={{fontSize:14,color:"#c9a84c",fontWeight:"bold"}}>{formatCurrency(calc.cajaTeor)}</div>
          <div style={{fontSize:8,color:"#4a4335",marginTop:2}}>= Inicial + Ventas + Dep. − Gas. − Ret.</div>
        </div>
        <div style={{textAlign:"center"}}>
          <div style={{fontSize:9,color:"#5a5240",textTransform:"uppercase",letterSpacing:1,marginBottom:3}}>Caja Real</div>
          {cajaReal!==null && !editando
            ? <div style={{fontSize:14,color:"#d4c89a",fontWeight:"bold",cursor:editable?"pointer":"default"}} onClick={()=>editable&&setEditando(true)}>{formatCurrency(cajaReal)}</div>
            : editando
              ? <div style={{display:"flex",gap:4,justifyContent:"center"}}>
                  <input type="number" value={inputVal} onChange={e=>setInputVal(e.target.value)} onKeyDown={e=>e.key==="Enter"&&guardar()} style={{...inp,width:80,padding:"4px 6px",fontSize:11}} autoFocus/>
                  <button onClick={guardar} style={{background:"#2a5c3f",border:"none",color:"#4caf82",borderRadius:6,padding:"3px 8px",cursor:"pointer",fontSize:12}}>✓</button>
                </div>
              : editable
                ? <button onClick={()=>setEditando(true)} style={{background:"transparent",border:"1px dashed #3a3520",color:"#5a5240",borderRadius:6,padding:"4px 10px",cursor:"pointer",fontSize:10}}>+ Ingresar</button>
                : <div style={{fontSize:12,color:"#4a4335"}}>–</div>
          }
        </div>
      </div>

      {/* Diferencia */}
      {dif!==null && (
        <div style={{borderTop:"1px solid #2e2b22",paddingTop:10,display:"flex",alignItems:"center",justifyContent:"space-between"}}>
          <div style={{fontSize:11,color:hayDif?"#c0503a":"#4caf82",fontWeight:"bold"}}>
            {hayDif ? "⚠ DESCUADRE" : "✓ CUADRE EXACTO"}
          </div>
          <div style={{fontSize:16,color:hayDif?"#c0503a":"#4caf82",fontWeight:"bold"}}>
            {dif>0?"+":""}{formatCurrency(dif)}
          </div>
        </div>
      )}
      {dif!==null && hayDif && (
        <div style={{marginTop:6,fontSize:10,color:"#8a3a2a",background:"#2a0e0e",borderRadius:6,padding:"6px 10px"}}>
          {dif>0
            ? `Hay ${formatCurrency(Math.abs(dif))} más de lo esperado en caja.`
            : `Faltan ${formatCurrency(Math.abs(dif))} en caja respecto al cálculo teórico.`}
        </div>
      )}
    </div>
  );
}

// ─────────────────────────────────────────────
// RESUMEN DIARIO (modal/panel)
// ─────────────────────────────────────────────
function ResumenDiario({localId,dateKey,onClose}) {
  const data=loadDay(localId,dateKey);
  const c=calcDay(data);
  const nombre=localStorage.getItem(`nombre_${localId}`)||localId;

  return(
    <div style={{position:"fixed",inset:0,background:"rgba(0,0,0,0.8)",display:"flex",alignItems:"center",justifyContent:"center",zIndex:300}}>
      <div style={{background:"#1a1710",border:"1px solid #3a3520",borderRadius:16,padding:24,width:380,maxWidth:"92vw",maxHeight:"90vh",overflowY:"auto"}}>
        <div style={{display:"flex",justifyContent:"space-between",alignItems:"center",marginBottom:16}}>
          <div>
            <div style={{fontSize:10,letterSpacing:2,color:"#6a6047",textTransform:"uppercase"}}>Resumen Diario</div>
            <div style={{fontSize:14,color:"#d4c89a"}}>{nombre} · {formatDate(dateKey)}</div>
            <div style={{fontSize:10,color:"#6a6047"}}>Semana {getWeekNumber(dateKey)}</div>
          </div>
          <button onClick={onClose} style={{background:"transparent",border:"none",color:"#6a6047",fontSize:20,cursor:"pointer"}}>×</button>
        </div>

        {/* Fórmula */}
        <div style={{background:"#0f0e0b",border:"1px solid #2e2b22",borderRadius:10,padding:14,marginBottom:14}}>
          <div style={{fontSize:9,color:"#5a5240",letterSpacing:2,textTransform:"uppercase",marginBottom:10}}>Cálculo</div>
          {[
            ["Caja Inicial",data.saldoInicial,"#8a7f5e",""],
            ["+ Ventas efectivo",c.ventas,"#4caf82","+"],
            ["+ Depósitos",c.depositos,"#6ac8d8","+"],
            ["− Gastos",c.gastos,"#c0503a","−"],
            ["− Retiros",c.retiros,"#d48a3a","−"],
          ].map(([label,val,color,sign])=>(
            <div key={label} style={{display:"flex",justifyContent:"space-between",marginBottom:5,paddingBottom:5,borderBottom:"1px solid #1e1c14"}}>
              <span style={{fontSize:11,color:"#8a7f5e"}}>{label}</span>
              <span style={{fontSize:12,color,fontWeight:"bold"}}>{sign}{formatCurrency(val)}</span>
            </div>
          ))}
          <div style={{display:"flex",justifyContent:"space-between",paddingTop:4}}>
            <span style={{fontSize:12,color:"#d4c89a",fontWeight:"bold"}}>= Caja Teórica</span>
            <span style={{fontSize:15,color:"#c9a84c",fontWeight:"bold"}}>{formatCurrency(c.cajaTeor)}</span>
          </div>
        </div>

        {/* Arqueo */}
        {c.cajaReal!==null && (
          <div style={{background:c.diferencia!==0?"#1f0e0e":"#0e1f16",border:`1px solid ${c.diferencia!==0?"#5c2a2a":"#2a5c3f"}`,borderRadius:10,padding:14,marginBottom:14}}>
            <div style={{display:"flex",justifyContent:"space-between",marginBottom:6}}>
              <span style={{fontSize:11,color:"#8a7f5e"}}>Caja Real</span>
              <span style={{fontSize:13,color:"#d4c89a",fontWeight:"bold"}}>{formatCurrency(c.cajaReal)}</span>
            </div>
            <div style={{display:"flex",justifyContent:"space-between"}}>
              <span style={{fontSize:11,color:"#8a7f5e"}}>Diferencia</span>
              <span style={{fontSize:14,color:c.diferencia!==0?"#c0503a":"#4caf82",fontWeight:"bold"}}>
                {c.diferencia>0?"+":""}{formatCurrency(c.diferencia)} {c.diferencia===0?"✓":"⚠"}
              </span>
            </div>
          </div>
        )}

        {/* Movimientos por tipo */}
        {Object.entries(TIPOS_MOV).map(([tipo,info])=>{
          const movs=data.movimientos.filter(m=>m.tipo===tipo);
          if(!movs.length) return null;
          return(
            <div key={tipo} style={{marginBottom:10}}>
              <div style={{fontSize:9,color:info.color,letterSpacing:2,textTransform:"uppercase",marginBottom:6}}>{info.emoji} {info.label}</div>
              {movs.map(m=>(
                <div key={m.id} style={{display:"flex",justifyContent:"space-between",padding:"5px 0",borderBottom:"1px solid #1e1c14"}}>
                  <span style={{fontSize:11,color:"#8a7f5e"}}>{m.descripcion}</span>
                  <span style={{fontSize:11,color:info.color,fontWeight:"bold"}}>{formatCurrency(m.monto)}</span>
                </div>
              ))}
            </div>
          );
        })}

        {data.nota&&<div style={{background:"#110f07",border:"1px solid #2e2800",borderRadius:8,padding:10,marginTop:8}}>
          <div style={{fontSize:9,color:"#5a5240",textTransform:"uppercase",letterSpacing:2,marginBottom:4}}>Nota</div>
          <div style={{fontSize:11,color:"#c4b880",fontStyle:"italic"}}>{data.nota}</div>
        </div>}
      </div>
    </div>
  );
}

// ─────────────────────────────────────────────
// CALENDARIO
// ─────────────────────────────────────────────
function CalendarioMes({localId,onSelectDate,selectedDate,accent}) {
  const now=new Date();
  const [year,setYear]=useState(now.getFullYear());
  const [month,setMonth]=useState(now.getMonth()+1);
  const today=getDateKey();
  const allReg=new Set(getAllKeys(localId));
  const days=getDaysInMonth(year,month);
  const weeks=groupByWeek(days);
  const weekNums=Object.keys(weeks).map(Number).sort((a,b)=>a-b);

  function prevMonth(){if(month===1){setYear(y=>y-1);setMonth(12);}else setMonth(m=>m-1);}
  function nextMonth(){
    const ny=month===12?year+1:year, nm=month===12?1:month+1;
    if(ny>now.getFullYear()||(ny===now.getFullYear()&&nm>now.getMonth()+1)) return;
    setYear(ny);setMonth(nm);
  }

  return(
    <div style={{background:"#0f0e0b",border:"1px solid #2e2b22",borderRadius:12,padding:14,marginBottom:14}}>
      <div style={{display:"flex",alignItems:"center",justifyContent:"space-between",marginBottom:12}}>
        <button onClick={prevMonth} style={{background:"transparent",border:"none",color:"#6a6047",cursor:"pointer",fontSize:16,padding:"0 6px"}}>‹</button>
        <div style={{fontSize:13,color:"#d4c89a",fontWeight:"bold"}}>{MESES[month-1]} {year}</div>
        <button onClick={nextMonth} style={{background:"transparent",border:"none",color:month===now.getMonth()+1&&year===now.getFullYear()?"#2e2b22":"#6a6047",cursor:"pointer",fontSize:16,padding:"0 6px"}}>›</button>
      </div>
      <div style={{display:"grid",gridTemplateColumns:"44px repeat(7,1fr)",gap:2,marginBottom:4}}>
        <div style={{fontSize:9,color:"#4a4335",textAlign:"center"}}>Sem.</div>
        {DIAS_CORTOS.map(d=><div key={d} style={{fontSize:9,color:"#4a4335",textAlign:"center"}}>{d}</div>)}
      </div>
      {weekNums.map(wNum=>{
        const wDays=weeks[wNum];
        let wIng=0,wEgr=0;
        wDays.forEach(dk=>{const d=loadDay(localId,dk);const c=calcDay(d);wIng+=c.ventas+c.depositos;wEgr+=c.gastos+c.retiros;});
        const cells=Array(7).fill(null);
        wDays.forEach(dk=>{const dow=new Date(...dk.split("-").map((v,i)=>i===1?+v-1:+v)).getDay();cells[dow]=dk;});
        return(
          <div key={wNum} style={{marginBottom:5}}>
            <div style={{display:"grid",gridTemplateColumns:"44px repeat(7,1fr)",gap:2,alignItems:"center"}}>
              <div style={{fontSize:9,color:accent+"cc",textAlign:"center",background:accent+"18",borderRadius:6,padding:"3px 0",fontWeight:"bold"}}>S{wNum}</div>
              {cells.map((dk,i)=>{
                if(!dk) return <div key={i}/>;
                const isFut=dk>today, hasDat=allReg.has(dk), isSel=dk===selectedDate, isTod=dk===today;
                const d=hasDat?loadDay(localId,dk):null;
                const hasDif=d&&calcDay(d).diferencia!==null&&calcDay(d).diferencia!==0;
                return(
                  <button key={dk} onClick={()=>!isFut&&onSelectDate(dk)}
                    style={{padding:"5px 2px",borderRadius:7,border:isSel?`2px solid ${accent}`:isTod?"2px solid #4caf8266":"1px solid transparent",
                      background:isSel?accent+"33":hasDat?"#1e2a1e":"#151310",
                      color:isFut?"#2a2820":isTod?"#4caf82":hasDat?"#c8e0c8":"#6a6047",
                      cursor:isFut?"default":"pointer",fontSize:11,textAlign:"center",position:"relative"}}>
                    {dk.split("-")[2]}
                    {hasDat&&<span style={{position:"absolute",bottom:2,left:"50%",transform:"translateX(-50%)",width:4,height:4,borderRadius:"50%",background:hasDif?"#c0503a":"#4caf82",display:"block"}}/>}
                  </button>);
              })}
            </div>
            {(wIng>0||wEgr>0)&&<div style={{display:"grid",gridTemplateColumns:"44px 1fr",gap:2,marginTop:2}}>
              <div/>
              <div style={{fontSize:9,color:"#5a5240",paddingLeft:4}}>
                <span style={{color:"#4caf8288"}}>+{formatCurrencyShort(wIng)}</span>
                <span style={{color:"#5a5240",margin:"0 4px"}}>·</span>
                <span style={{color:"#c0503a88"}}>-{formatCurrencyShort(wEgr)}</span>
              </div>
            </div>}
          </div>);
      })}
    </div>);
}

// ─────────────────────────────────────────────
// INFORMES
// ─────────────────────────────────────────────
function Informes() {
  const now=new Date();
  const today=getDateKey();
  const [tipoInforme,setTipoInforme]=useState("diario");
  const [year,setYear]=useState(now.getFullYear());
  const [month,setMonth]=useState(now.getMonth()+1);
  const [weekNum,setWeekNum]=useState(getWeekNumber(today));
  const years=[]; for(let y=now.getFullYear();y>=now.getFullYear()-2;y--) years.push(y);

  // Helpers
  function getKeysForMonth(localId,y,m) {
    const ms=`${y}-${String(m).padStart(2,"0")}`;
    return getAllKeys(localId).filter(k=>k.startsWith(ms)).sort();
  }
  function getKeysForWeek(localId,y,wn) {
    return getAllKeys(localId).filter(k=>k.startsWith(String(y))&&getWeekNumber(k)===wn).sort();
  }
  function sumKeys(localId,keys) {
    let ventas=0,depositos=0,gastos=0,retiros=0,descuadres=0;
    keys.forEach(dk=>{const d=loadDay(localId,dk);const c=calcDay(d);ventas+=c.ventas;depositos+=c.depositos;gastos+=c.gastos;retiros+=c.retiros;if(c.diferencia!==null&&c.diferencia!==0)descuadres++;});
    return {ventas,depositos,gastos,retiros,utilidad:ventas+depositos-gastos-retiros,descuadres,dias:keys.length};
  }

  const nombre1=localStorage.getItem("nombre_local1")||"Cornella";
  const nombre2=localStorage.getItem("nombre_local2")||"Badalona";

  // ── Informe diario ──
  function InfoDiario() {
    const [fecha,setFecha]=useState(today);
    return(
      <div>
        <div style={{display:"flex",gap:8,alignItems:"center",marginBottom:16,flexWrap:"wrap"}}>
          <input type="date" value={fecha} max={today} onChange={e=>setFecha(e.target.value)}
            style={{...inp,fontSize:12}}/>
          <span style={{fontSize:10,color:"#6a6047"}}>Sem. {getWeekNumber(fecha)}</span>
        </div>
        <div style={{display:"grid",gridTemplateColumns:"1fr 1fr",gap:14}}>
          {LOCALES.map(local=>{
            const nombre=localStorage.getItem(`nombre_${local.id}`)||local.nombre;
            const d=loadDay(local.id,fecha);
            const c=calcDay(d);
            return(
              <div key={local.id} style={{background:"#0f0e0b",border:`1px solid ${local.color}33`,borderRadius:12,padding:14}}>
                <div style={{fontSize:11,color:local.color,marginBottom:10,fontWeight:"bold"}}>{local.emoji} {nombre}</div>
                {[
                  ["Caja Inicial",d.saldoInicial,"#8a7f5e"],
                  ["Ventas",c.ventas,"#4caf82"],
                  ["Depósitos",c.depositos,"#6ac8d8"],
                  ["Gastos",c.gastos,"#c0503a"],
                  ["Retiros",c.retiros,"#d48a3a"],
                ].map(([l,v,col])=>(
                  <div key={l} style={{display:"flex",justifyContent:"space-between",marginBottom:5}}>
                    <span style={{fontSize:10,color:"#5a5240"}}>{l}</span>
                    <span style={{fontSize:11,color:col,fontWeight:"bold"}}>{formatCurrency(v)}</span>
                  </div>
                ))}
                <div style={{borderTop:"1px solid #2e2b22",paddingTop:6,display:"flex",justifyContent:"space-between"}}>
                  <span style={{fontSize:11,color:"#c9a84c"}}>Caja Teórica</span>
                  <span style={{fontSize:13,color:"#c9a84c",fontWeight:"bold"}}>{formatCurrency(c.cajaTeor)}</span>
                </div>
                {c.cajaReal!==null&&<div style={{display:"flex",justifyContent:"space-between",marginTop:4}}>
                  <span style={{fontSize:10,color:c.diferencia!==0?"#c0503a":"#4caf82"}}>{c.diferencia!==0?"⚠ Descuadre":"✓ Cuadre"}</span>
                  <span style={{fontSize:11,color:c.diferencia!==0?"#c0503a":"#4caf82",fontWeight:"bold"}}>{c.diferencia>0?"+":""}{formatCurrency(c.diferencia)}</span>
                </div>}
                <div style={{fontSize:10,color:d.cerrado?"#4caf82":"#8a7f5e",marginTop:6}}>{d.cerrado?"Cerrada ✓":"Abierta"}</div>
              </div>
            );
          })}
        </div>
      </div>
    );
  }

  // ── Informe semanal ──
  function InfoSemanal() {
    const weeks=[];
    for(let w=1;w<=53;w++) weeks.push(w);
    const keys1=getKeysForWeek("local1",year,weekNum);
    const keys2=getKeysForWeek("local2",year,weekNum);
    const s1=sumKeys("local1",keys1);
    const s2=sumKeys("local2",keys2);

    const dias=[...new Set([...keys1,...keys2])].sort().map(dk=>{
      const c1=calcDay(loadDay("local1",dk));const c2=calcDay(loadDay("local2",dk));
      return{dia:formatDateShort(dk),[`${nombre1}`]:c1.ventas+c1.depositos,[`${nombre2}`]:c2.ventas+c2.depositos};
    });

    return(
      <div>
        <div style={{display:"flex",gap:8,marginBottom:16,flexWrap:"wrap"}}>
          <select value={year} onChange={e=>setYear(+e.target.value)} style={{...inp,fontSize:12}}>
            {years.map(y=><option key={y} value={y}>{y}</option>)}
          </select>
          <select value={weekNum} onChange={e=>setWeekNum(+e.target.value)} style={{...inp,fontSize:12}}>
            {weeks.map(w=><option key={w} value={w}>Semana {w}</option>)}
          </select>
        </div>
        <ComparativaCards s1={s1} s2={s2} n1={nombre1} n2={nombre2} c1={LOCALES[0].color} c2={LOCALES[1].color}/>
        {dias.length>0&&<>
          <div style={{fontSize:10,color:"#6a6047",letterSpacing:2,textTransform:"uppercase",margin:"16px 0 8px"}}>Ingresos diarios</div>
          <ResponsiveContainer width="100%" height={160}>
            <BarChart data={dias}><CartesianGrid strokeDasharray="3 3" stroke="#2a2720"/>
              <XAxis dataKey="dia" tick={{fill:"#6a6047",fontSize:9}} axisLine={false} tickLine={false}/>
              <YAxis tick={{fill:"#6a6047",fontSize:9}} tickFormatter={formatCurrencyShort} axisLine={false} tickLine={false} width={38}/>
              <Tooltip content={<CustomTooltip/>}/><Legend wrapperStyle={{fontSize:10}}/>
              <Bar dataKey={nombre1} fill="#c9a84c" radius={[3,3,0,0]}/><Bar dataKey={nombre2} fill="#6a9fd8" radius={[3,3,0,0]}/>
            </BarChart>
          </ResponsiveContainer>
        </>}
      </div>
    );
  }

  // ── Informe mensual ──
  function InfoMensual() {
    const keys1=getKeysForMonth("local1",year,month);
    const keys2=getKeysForMonth("local2",year,month);
    const s1=sumKeys("local1",keys1);
    const s2=sumKeys("local2",keys2);
    const allDays=[...new Set([...keys1,...keys2])].sort();
    const chartData=allDays.map(dk=>{
      const c1=calcDay(loadDay("local1",dk));const c2=calcDay(loadDay("local2",dk));
      return{dia:dk.split("-")[2],[`${nombre1}`]:c1.ventas+c1.depositos,[`${nombre2}`]:c2.ventas+c2.depositos,
        [`${nombre1} Gas`]:c1.gastos+c1.retiros,[`${nombre2} Gas`]:c2.gastos+c2.retiros};
    });
    return(
      <div>
        <div style={{display:"flex",gap:8,marginBottom:16,flexWrap:"wrap"}}>
          <select value={year} onChange={e=>setYear(+e.target.value)} style={{...inp,fontSize:12}}>{years.map(y=><option key={y}>{y}</option>)}</select>
          <select value={month} onChange={e=>setMonth(+e.target.value)} style={{...inp,fontSize:12}}>{MESES.map((m,i)=><option key={i} value={i+1}>{m}</option>)}</select>
        </div>
        <ComparativaCards s1={s1} s2={s2} n1={nombre1} n2={nombre2} c1={LOCALES[0].color} c2={LOCALES[1].color}/>
        {chartData.length>0&&<>
          <div style={{fontSize:10,color:"#6a6047",letterSpacing:2,textTransform:"uppercase",margin:"16px 0 8px"}}>Ingresos del mes</div>
          <ResponsiveContainer width="100%" height={160}>
            <BarChart data={chartData}><CartesianGrid strokeDasharray="3 3" stroke="#2a2720"/>
              <XAxis dataKey="dia" tick={{fill:"#6a6047",fontSize:9}} axisLine={false} tickLine={false}/>
              <YAxis tick={{fill:"#6a6047",fontSize:9}} tickFormatter={formatCurrencyShort} axisLine={false} tickLine={false} width={38}/>
              <Tooltip content={<CustomTooltip/>}/><Legend wrapperStyle={{fontSize:10}}/>
              <Bar dataKey={nombre1} fill="#c9a84c" radius={[3,3,0,0]}/><Bar dataKey={nombre2} fill="#6a9fd8" radius={[3,3,0,0]}/>
            </BarChart>
          </ResponsiveContainer>
          <div style={{fontSize:10,color:"#6a6047",letterSpacing:2,textTransform:"uppercase",margin:"16px 0 8px"}}>Gastos del mes</div>
          <ResponsiveContainer width="100%" height={160}>
            <BarChart data={chartData}><CartesianGrid strokeDasharray="3 3" stroke="#2a2720"/>
              <XAxis dataKey="dia" tick={{fill:"#6a6047",fontSize:9}} axisLine={false} tickLine={false}/>
              <YAxis tick={{fill:"#6a6047",fontSize:9}} tickFormatter={formatCurrencyShort} axisLine={false} tickLine={false} width={38}/>
              <Tooltip content={<CustomTooltip/>}/><Legend wrapperStyle={{fontSize:10}}/>
              <Bar dataKey={`${nombre1} Gas`} fill="#8a3a2a" radius={[3,3,0,0]}/><Bar dataKey={`${nombre2} Gas`} fill="#3a5a8a" radius={[3,3,0,0]}/>
            </BarChart>
          </ResponsiveContainer>
        </>}
      </div>
    );
  }

  // ── Comparativa ──
  function InfoComparativa() {
    const mesesData=Array.from({length:12},(_,i)=>{
      const m=i+1;
      const k1=getKeysForMonth("local1",year,m);
      const k2=getKeysForMonth("local2",year,m);
      const s1=sumKeys("local1",k1);
      const s2=sumKeys("local2",k2);
      return{mes:MESES[i].slice(0,3),[nombre1]:s1.utilidad,[nombre2]:s2.utilidad,
        [`${nombre1}V`]:s1.ventas+s1.depositos,[`${nombre2}V`]:s2.ventas+s2.depositos};
    });
    return(
      <div>
        <div style={{display:"flex",gap:8,marginBottom:16}}>
          <select value={year} onChange={e=>setYear(+e.target.value)} style={{...inp,fontSize:12}}>{years.map(y=><option key={y}>{y}</option>)}</select>
        </div>
        <div style={{fontSize:10,color:"#6a6047",letterSpacing:2,textTransform:"uppercase",marginBottom:8}}>Utilidad mensual comparada</div>
        <ResponsiveContainer width="100%" height={200}>
          <BarChart data={mesesData}><CartesianGrid strokeDasharray="3 3" stroke="#2a2720"/>
            <XAxis dataKey="mes" tick={{fill:"#6a6047",fontSize:10}} axisLine={false} tickLine={false}/>
            <YAxis tick={{fill:"#6a6047",fontSize:9}} tickFormatter={formatCurrencyShort} axisLine={false} tickLine={false} width={40}/>
            <Tooltip content={<CustomTooltip/>}/><Legend wrapperStyle={{fontSize:10}}/>
            <Bar dataKey={nombre1} fill="#c9a84c" radius={[3,3,0,0]}/><Bar dataKey={nombre2} fill="#6a9fd8" radius={[3,3,0,0]}/>
          </BarChart>
        </ResponsiveContainer>
        <div style={{fontSize:10,color:"#6a6047",letterSpacing:2,textTransform:"uppercase",margin:"18px 0 8px"}}>Ventas + Depósitos mensuales</div>
        <ResponsiveContainer width="100%" height={200}>
          <LineChart data={mesesData}><CartesianGrid strokeDasharray="3 3" stroke="#2a2720"/>
            <XAxis dataKey="mes" tick={{fill:"#6a6047",fontSize:10}} axisLine={false} tickLine={false}/>
            <YAxis tick={{fill:"#6a6047",fontSize:9}} tickFormatter={formatCurrencyShort} axisLine={false} tickLine={false} width={40}/>
            <Tooltip content={<CustomTooltip/>}/><Legend wrapperStyle={{fontSize:10}}/>
            <Line type="monotone" dataKey={`${nombre1}V`} name={nombre1} stroke="#c9a84c" strokeWidth={2} dot={{r:3}}/>
            <Line type="monotone" dataKey={`${nombre2}V`} name={nombre2} stroke="#6a9fd8" strokeWidth={2} dot={{r:3}}/>
          </LineChart>
        </ResponsiveContainer>
      </div>
    );
  }

  return(
    <div style={{background:"#141210",border:"1px solid #2e2b22",borderRadius:16,padding:20,marginBottom:24}}>
      <div style={{fontSize:11,letterSpacing:2,color:"#8a7f5e",textTransform:"uppercase",marginBottom:14}}>📋 Informes</div>
      <div style={{display:"flex",gap:6,marginBottom:20,flexWrap:"wrap"}}>
        {[["diario","📅 Diario"],["semanal","📆 Semanal"],["mensual","🗓 Mensual"],["comparativa","⚖ Comparativa"]].map(([k,label])=>(
          <button key={k} onClick={()=>setTipoInforme(k)}
            style={{padding:"6px 14px",borderRadius:20,border:`1px solid ${tipoInforme===k?"#c9a84c":"#2e2b22"}`,background:tipoInforme===k?"#2a2010":"transparent",color:tipoInforme===k?"#c9a84c":"#6a6047",fontSize:11,cursor:"pointer"}}>
            {label}
          </button>
        ))}
      </div>
      {tipoInforme==="diario"&&<InfoDiario/>}
      {tipoInforme==="semanal"&&<InfoSemanal/>}
      {tipoInforme==="mensual"&&<InfoMensual/>}
      {tipoInforme==="comparativa"&&<InfoComparativa/>}
    </div>
  );
}

function ComparativaCards({s1,s2,n1,n2,c1,c2}) {
  return(
    <div style={{display:"grid",gridTemplateColumns:"1fr 1fr",gap:12,marginBottom:10}}>
      {[[s1,n1,c1],[s2,n2,c2]].map(([s,n,c])=>(
        <div key={n} style={{background:"#0f0e0b",border:`1px solid ${c}33`,borderRadius:10,padding:12}}>
          <div style={{fontSize:10,color:c,marginBottom:8,fontWeight:"bold"}}>{n}</div>
          {[["Ventas",s.ventas,"#4caf82"],["Depósitos",s.depositos,"#6ac8d8"],["Gastos",s.gastos,"#c0503a"],["Retiros",s.retiros,"#d48a3a"]].map(([l,v,col])=>(
            <div key={l} style={{display:"flex",justifyContent:"space-between",marginBottom:3}}>
              <span style={{fontSize:10,color:"#5a5240"}}>{l}</span>
              <span style={{fontSize:10,color:col,fontWeight:"bold"}}>{formatCurrency(v)}</span>
            </div>
          ))}
          <div style={{borderTop:"1px solid #2e2b22",paddingTop:5,display:"flex",justifyContent:"space-between"}}>
            <span style={{fontSize:10,color:"#8a7f5e"}}>Utilidad</span>
            <span style={{fontSize:13,color:s.utilidad>=0?"#c9a84c":"#c0503a",fontWeight:"bold"}}>{formatCurrency(s.utilidad)}</span>
          </div>
          {s.descuadres>0&&<div style={{fontSize:9,color:"#c0503a",marginTop:4}}>⚠ {s.descuadres} descuadre{s.descuadres>1?"s":""} en el período</div>}
        </div>
      ))}
    </div>
  );
}

// ─────────────────────────────────────────────
// CAJA LOCAL
// ─────────────────────────────────────────────
function CajaLocal({local}) {
  const today=getDateKey();
  const [viewDate,setViewDate]=useState(today);
  const [tab,setTab]=useState("hoy");
  const [dayData,setDayData]=useState(()=>loadDay(local.id,today));
  const [form,setForm]=useState({tipo:"venta",categoria:CATS.venta[0],descripcion:"",monto:""});
  const [editSaldo,setEditSaldo]=useState(false);
  const [saldoInput,setSaldoInput]=useState("");
  const [showCierre,setShowCierre]=useState(false);
  const [showResumen,setShowResumen]=useState(false);
  const [nota,setNota]=useState("");
  const [flash,setFlash]=useState("");
  const [localNombre,setLocalNombre]=useState(()=>localStorage.getItem(`nombre_${local.id}`)||local.nombre);
  const inputRef=useRef();

  useEffect(()=>{const d=loadDay(local.id,viewDate);setDayData(d);setNota(d.nota||"");},[viewDate,local.id]);

  function persist(updated){saveDay(local.id,viewDate,updated);setDayData(updated);}
  function showFlash(msg){setFlash(msg);setTimeout(()=>setFlash(""),2000);}

  function addMovimiento(){
    if(!form.monto||isNaN(+form.monto)||+form.monto<=0) return;
    const mov={id:Date.now(),tipo:form.tipo,categoria:form.categoria,descripcion:form.descripcion.trim()||form.categoria,monto:+form.monto,hora:new Date().toLocaleTimeString("es-ES",{hour:"2-digit",minute:"2-digit"})};
    persist({...dayData,movimientos:[...dayData.movimientos,mov]});
    setForm(f=>({...f,descripcion:"",monto:""}));
    showFlash(`✓ ${TIPOS_MOV[form.tipo].label} registrado`);
    setTimeout(()=>inputRef.current?.focus(),50);
  }

  function deleteMovimiento(id){persist({...dayData,movimientos:dayData.movimientos.filter(m=>m.id!==id)});}
  function guardarSaldo(){const v=parseFloat(saldoInput);if(!isNaN(v))persist({...dayData,saldoInicial:v});setEditSaldo(false);}
  function cerrarCaja(){persist({...dayData,cerrado:true,nota,horaCierre:new Date().toLocaleTimeString("es-ES")});setShowCierre(false);showFlash("✓ Caja cerrada");}

  const calc=calcDay(dayData);
  const isToday=viewDate===today;
  const isFuture=viewDate>today;
  const accent=local.color;
  const semana=getWeekNumber(viewDate);
  const hayDescuadre=calc.diferencia!==null&&calc.diferencia!==0;

  return(
    <div style={{background:"#141210",border:`1px solid ${hayDescuadre?"#5c2a2a":accent+"33"}`,borderRadius:16,overflow:"hidden",position:"relative"}}>
      {flash&&<div style={{position:"absolute",top:12,left:"50%",transform:"translateX(-50%)",background:"#2a5c3f",color:"#b8f5c8",padding:"8px 22px",borderRadius:40,fontSize:12,zIndex:99,boxShadow:"0 4px 16px rgba(0,0,0,0.5)",whiteSpace:"nowrap"}}>{flash}</div>}

      {/* Header */}
      <div style={{background:`${accent}18`,borderBottom:`1px solid ${accent}33`,padding:"12px 16px",display:"flex",alignItems:"center",justifyContent:"space-between"}}>
        <div style={{display:"flex",alignItems:"center",gap:10}}>
          <span style={{fontSize:20}}>{local.emoji}</span>
          <div>
            <div style={{fontSize:10,letterSpacing:2,color:accent+"aa",textTransform:"uppercase"}}>Local</div>
            <div style={{fontSize:15,color:accent,fontWeight:"bold"}}><LocalNameEditor local={local} onSave={setLocalNombre}/></div>
          </div>
        </div>
        <div style={{textAlign:"right"}}>
          <div style={{fontSize:19,color:calc.cajaTeor>=0?"#4caf82":"#c0503a",fontWeight:"bold"}}>{formatCurrency(calc.cajaTeor)}</div>
          <div style={{fontSize:9,color:"#5a5240",letterSpacing:1}}>CAJA TEÓRICA · {isToday?"HOY":formatDateShort(viewDate).toUpperCase()}</div>
          {hayDescuadre&&<div style={{fontSize:10,color:"#c0503a"}}>⚠ Descuadre {formatCurrency(calc.diferencia)}</div>}
        </div>
      </div>

      {/* Sub-tabs */}
      <div style={{display:"flex",borderBottom:"1px solid #2e2b22"}}>
        {[["hoy","📋 Caja"],["historial","📅 Calendario"]].map(([k,label])=>(
          <button key={k} onClick={()=>{setTab(k);if(k==="hoy")setViewDate(today);}}
            style={{flex:1,padding:"9px 0",background:"transparent",border:"none",borderBottom:tab===k?`2px solid ${accent}`:"2px solid transparent",color:tab===k?accent:"#6a6047",fontSize:11,cursor:"pointer",letterSpacing:1}}>
            {label}
          </button>
        ))}
      </div>

      <div style={{padding:14}}>
        {/* Semana + fecha */}
        <div style={{display:"flex",alignItems:"center",gap:8,marginBottom:12,flexWrap:"wrap"}}>
          <span style={{background:accent+"22",border:`1px solid ${accent}44`,color:accent,borderRadius:20,padding:"3px 10px",fontSize:10,letterSpacing:1,fontWeight:"bold"}}>Semana {semana}</span>
          <span style={{fontSize:11,color:"#5a5240"}}>{formatDateShort(viewDate)}{isToday?" · Hoy":""}</span>
          {dayData.cerrado&&<span style={{fontSize:10,color:accent}}>· Cerrada ✓</span>}
          <button onClick={()=>setShowResumen(true)} style={{marginLeft:"auto",background:"transparent",border:"1px solid #2e2b22",color:"#6a6047",padding:"3px 10px",borderRadius:20,fontSize:10,cursor:"pointer"}}>
            📄 Ver resumen
          </button>
        </div>

        {/* 4 métricas */}
        <div style={{display:"grid",gridTemplateColumns:"1fr 1fr 1fr 1fr",gap:6,marginBottom:12}}>
          {[["Ventas",calc.ventas,"#4caf82"],["Depósitos",calc.depositos,"#6ac8d8"],["Gastos",calc.gastos,"#c0503a"],["Retiros",calc.retiros,"#d48a3a"]].map(([label,val,color])=>(
            <div key={label} style={{background:"#0f0e0b",border:"1px solid #2e2b22",borderRadius:8,padding:"8px 6px",textAlign:"center"}}>
              <div style={{fontSize:8,letterSpacing:1,color:"#4a4335",textTransform:"uppercase",marginBottom:3}}>{label}</div>
              <div style={{fontSize:11,color,fontWeight:"bold"}}>{formatCurrency(val)}</div>
            </div>
          ))}
        </div>

        {/* Calendario */}
        {tab==="historial"&&<CalendarioMes localId={local.id} onSelectDate={setViewDate} selectedDate={viewDate} accent={accent}/>}

        {/* Arqueo */}
        {!isFuture&&(
          <DescuadrePanel
            calc={calc}
            cajaReal={dayData.cajaReal??null}
            onSetCajaReal={v=>persist({...dayData,cajaReal:v})}
            editable={!dayData.cerrado}
          />
        )}

        {/* Acciones */}
        {!isFuture&&!dayData.cerrado&&(
          <div style={{display:"flex",gap:6,marginBottom:12}}>
            <button onClick={()=>setEditSaldo(true)} style={{...btnSec,fontSize:10}}>✏️ Saldo inicial</button>
            <button onClick={()=>setShowCierre(true)} style={{background:accent,border:"none",color:"#0f0e0b",padding:"7px 14px",borderRadius:20,fontSize:11,fontWeight:"bold",cursor:"pointer"}}>🔒 Cerrar caja</button>
          </div>
        )}
        {dayData.cerrado&&<div style={{fontSize:10,color:accent,marginBottom:10}}>CERRADA ✓ {dayData.horaCierre} <span onClick={()=>persist({...dayData,cerrado:false})} style={{color:"#6a6047",cursor:"pointer",borderBottom:"1px dashed #4a4335",marginLeft:8}}>Reabrir</span></div>}

        {/* Saldo inicial */}
        {editSaldo&&(
          <div style={{background:"#0f0e0b",border:"1px solid #2e2b22",borderRadius:8,padding:12,marginBottom:12}}>
            <div style={{fontSize:11,color:"#6a6047",marginBottom:6}}>Saldo inicial (efectivo al abrir)</div>
            <div style={{display:"flex",gap:6}}>
              <input type="number" value={saldoInput} onChange={e=>setSaldoInput(e.target.value)} placeholder="0" style={{...inp,flex:1}}/>
              <button onClick={guardarSaldo} style={{...btnPri,background:accent}}>Guardar</button>
              <button onClick={()=>setEditSaldo(false)} style={btnSec}>×</button>
            </div>
          </div>
        )}

        {/* Formulario movimiento */}
        {!isFuture&&!dayData.cerrado&&(
          <div style={{background:"#0f0e0b",border:"1px solid #2e2b22",borderRadius:10,padding:14,marginBottom:12}}>
            {!isToday&&<div style={{fontSize:10,color:"#c9a84c",marginBottom:8}}>⚠ Registrando en día pasado: {formatDateShort(viewDate)}</div>}
            {/* Tipo */}
            <div style={{display:"grid",gridTemplateColumns:"1fr 1fr 1fr 1fr",gap:5,marginBottom:10}}>
              {Object.entries(TIPOS_MOV).map(([tipo,info])=>(
                <button key={tipo} onClick={()=>setForm(f=>({...f,tipo,categoria:CATS[tipo][0]}))}
                  style={{padding:"7px 4px",borderRadius:7,border:`1px solid ${form.tipo===tipo?info.color:"#2e2b22"}`,
                    background:form.tipo===tipo?info.color+"22":"transparent",color:form.tipo===tipo?info.color:"#5a5240",fontSize:10,cursor:"pointer",textAlign:"center"}}>
                  {info.emoji}<br/><span style={{fontSize:9}}>{info.label.split(" ")[0]}</span>
                </button>
              ))}
            </div>
            <select value={form.categoria} onChange={e=>setForm(f=>({...f,categoria:e.target.value}))} style={{...inp,width:"100%",boxSizing:"border-box",marginBottom:7}}>
              {CATS[form.tipo].map(c=><option key={c}>{c}</option>)}
            </select>
            <input type="text" placeholder="Descripción (opcional)" value={form.descripcion} onChange={e=>setForm(f=>({...f,descripcion:e.target.value}))} style={{...inp,width:"100%",boxSizing:"border-box",marginBottom:7}}/>
            <div style={{display:"flex",gap:6}}>
              <input ref={inputRef} type="number" placeholder="Monto" value={form.monto} onChange={e=>setForm(f=>({...f,monto:e.target.value}))} onKeyDown={e=>e.key==="Enter"&&addMovimiento()} style={{...inp,flex:1}}/>
              <button onClick={addMovimiento} style={{...btnPri,background:TIPOS_MOV[form.tipo].color}}>+ Agregar</button>
            </div>
          </div>
        )}

        {/* Modal cierre */}
        {showCierre&&(
          <div style={{position:"fixed",inset:0,background:"rgba(0,0,0,0.75)",display:"flex",alignItems:"center",justifyContent:"center",zIndex:200}}>
            <div style={{background:"#1a1710",border:`1px solid ${accent}44`,borderRadius:14,padding:24,width:320,maxWidth:"90vw"}}>
              <div style={{fontSize:14,color:accent,marginBottom:4}}>Cerrar Caja · {localNombre}</div>
              <div style={{fontSize:11,color:"#8a7f5e",marginBottom:4}}>{formatDateShort(viewDate)} · Semana {semana}</div>
              <div style={{fontSize:13,color:"#8a7f5e",marginBottom:4}}>Caja Teórica: <strong style={{color:"#c9a84c"}}>{formatCurrency(calc.cajaTeor)}</strong></div>
              {calc.cajaReal!==null&&<div style={{fontSize:12,color:hayDescuadre?"#c0503a":"#4caf82",marginBottom:10}}>
                {hayDescuadre?`⚠ Descuadre de ${formatCurrency(calc.diferencia)}`:"✓ Caja cuadrada"}
              </div>}
              <textarea placeholder="Notas del día (opcional)" value={nota} onChange={e=>setNota(e.target.value)} style={{...inp,width:"100%",boxSizing:"border-box",height:70,resize:"vertical",marginBottom:14}}/>
              <div style={{display:"flex",gap:8}}>
                <button onClick={cerrarCaja} style={{...btnPri,background:accent,flex:1}}>Confirmar</button>
                <button onClick={()=>setShowCierre(false)} style={{...btnSec,flex:1}}>Cancelar</button>
              </div>
            </div>
          </div>
        )}

        {/* Modal resumen diario */}
        {showResumen&&<ResumenDiario localId={local.id} dateKey={viewDate} onClose={()=>setShowResumen(false)}/>}

        {/* Lista movimientos */}
        <div>
          {dayData.movimientos.length===0&&!isFuture&&<div style={{textAlign:"center",color:"#3a3520",fontSize:12,padding:"14px 0"}}>Sin movimientos{isToday?" aún":""}</div>}
          {isFuture&&<div style={{textAlign:"center",color:"#3a3520",fontSize:12,padding:"14px 0"}}>Día futuro</div>}
          {dayData.movimientos.slice().reverse().map(mov=>{
            const info=TIPOS_MOV[mov.tipo]||{color:"#8a7f5e",signo:1};
            return(
              <div key={mov.id} style={{background:"#0a0908",border:"1px solid #222018",borderRadius:8,padding:"9px 11px",marginBottom:5,display:"flex",alignItems:"center",justifyContent:"space-between"}}>
                <div style={{display:"flex",alignItems:"center",gap:8,flex:1}}>
                  <div style={{width:7,height:7,borderRadius:"50%",background:info.color,flexShrink:0}}/>
                  <div>
                    <div style={{fontSize:11,color:"#c8c0aa"}}>{mov.descripcion}</div>
                    <div style={{fontSize:9,color:"#4a4335"}}>{mov.categoria} · {mov.hora}</div>
                  </div>
                </div>
                <div style={{display:"flex",alignItems:"center",gap:7}}>
                  <div style={{fontSize:12,fontWeight:"bold",color:info.color}}>{info.signo>0?"+":"-"}{formatCurrency(mov.monto)}</div>
                  {!dayData.cerrado&&<button onClick={()=>deleteMovimiento(mov.id)} style={{background:"none",border:"none",color:"#3a3520",cursor:"pointer",fontSize:15}}>×</button>}
                </div>
              </div>
            );
          })}
          {dayData.cerrado&&dayData.nota&&(
            <div style={{background:"#110f07",border:"1px solid #2e2800",borderRadius:8,padding:10,marginTop:8}}>
              <div style={{fontSize:9,letterSpacing:2,color:"#5a5240",textTransform:"uppercase",marginBottom:3}}>Nota</div>
              <div style={{fontSize:11,color:"#c4b880",fontStyle:"italic"}}>{dayData.nota}</div>
            </div>
          )}
        </div>
      </div>
    </div>
  );
}

// ─────────────────────────────────────────────
// RESUMEN CONSOLIDADO
// ─────────────────────────────────────────────
function ResumenConsolidado() {
  const today=getDateKey();
  const semana=getWeekNumber(today);
  const totales=LOCALES.map(local=>{
    const d=loadDay(local.id,today); const c=calcDay(d);
    return{local,calc:c,nombre:localStorage.getItem(`nombre_${local.id}`)||local.nombre};
  });
  const totalUtil=totales.reduce((s,t)=>s+t.calc.cajaTeor,0)-(loadDay("local1",today).saldoInicial||0)-(loadDay("local2",today).saldoInicial||0);
  const hayAlerta=totales.some(t=>t.calc.diferencia!==null&&t.calc.diferencia!==0);

  return(
    <div style={{background:"#1a1710",border:`1px solid ${hayAlerta?"#5c2a2a":"#3a3520"}`,borderRadius:14,padding:"14px 18px",marginBottom:16}}>
      <div style={{display:"flex",justifyContent:"space-between",alignItems:"center",marginBottom:10}}>
        <div style={{fontSize:10,letterSpacing:3,color:"#6a6047",textTransform:"uppercase"}}>Consolidado Hoy</div>
        <div style={{display:"flex",gap:8,alignItems:"center"}}>
          {hayAlerta&&<span style={{fontSize:10,color:"#c0503a",background:"#2a0e0e",border:"1px solid #5c2a2a",borderRadius:20,padding:"2px 8px"}}>⚠ Descuadre</span>}
          <span style={{background:"#2a2010",border:"1px solid #c9a84c44",color:"#c9a84c",borderRadius:20,padding:"3px 10px",fontSize:10,fontWeight:"bold"}}>Semana {semana}</span>
        </div>
      </div>
      <div style={{display:"grid",gridTemplateColumns:"1fr 1fr 1fr",gap:10,marginBottom:10}}>
        {[["Ventas+Dep.",totales.reduce((s,t)=>s+t.calc.ventas+t.calc.depositos,0),"#4caf82"],
          ["Gastos+Ret.",totales.reduce((s,t)=>s+t.calc.gastos+t.calc.retiros,0),"#c0503a"],
          ["Cajas Teór.",totales.reduce((s,t)=>s+t.calc.cajaTeor,0),totalUtil>=0?"#c9a84c":"#c0503a"]
        ].map(([label,val,color])=>(
          <div key={label} style={{textAlign:"center"}}>
            <div style={{fontSize:8,letterSpacing:1,color:"#5a5240",textTransform:"uppercase",marginBottom:3}}>{label}</div>
            <div style={{fontSize:14,color,fontWeight:"bold"}}>{formatCurrency(val)}</div>
          </div>
        ))}
      </div>
      <div style={{borderTop:"1px solid #2e2b22",paddingTop:8,display:"flex",gap:12}}>
        {totales.map(({local,calc,nombre})=>(
          <div key={local.id} style={{flex:1,textAlign:"center"}}>
            <div style={{fontSize:9,color:local.color,marginBottom:2}}>{nombre}</div>
            <div style={{fontSize:12,color:calc.cajaTeor>=0?"#4caf82":"#c0503a",fontWeight:"bold"}}>{formatCurrency(calc.cajaTeor)}</div>
            {calc.diferencia!==null&&calc.diferencia!==0&&<div style={{fontSize:9,color:"#c0503a"}}>⚠ {formatCurrency(calc.diferencia)}</div>}
          </div>
        ))}
      </div>
    </div>
  );
}

// ─────────────────────────────────────────────
// APP PRINCIPAL
// ─────────────────────────────────────────────
export default function App() {
  const [mainTab,setMainTab]=useState("caja");
  return(
    <div style={{fontFamily:"'Georgia', serif",minHeight:"100vh",background:"#0f0e0b",color:"#e8e0cc",padding:"0 0 60px"}}>
      <div style={{background:"#1a1710",borderBottom:"1px solid #2e2b22",padding:"10px 18px",display:"flex",justifyContent:"space-between",alignItems:"center"}}>
        <div style={{display:"flex",alignItems:"center",gap:12}}>
          <img src={require("./logo.png")} alt="Logo" style={{height:48,borderRadius:6,objectFit:"contain"}}/>
          <div>
            <div style={{fontSize:9,letterSpacing:3,color:"#6a6047",textTransform:"uppercase"}}>Libro de Caja</div>
            <div style={{fontSize:16,color:"#f0e8d0"}}>Mis Negocios</div>
          </div>
        </div>
        <div style={{display:"flex",alignItems:"center",gap:8}}>
          <button onClick={exportarExcel} style={{background:"#1e3a1e",border:"1px solid #2a5c2a",color:"#4caf82",padding:"6px 12px",borderRadius:20,fontSize:10,cursor:"pointer",fontFamily:"Georgia,serif"}}>📊 Excel</button>
          <div style={{fontSize:10,color:"#6a6047"}}>{new Date().toLocaleDateString("es-ES",{weekday:"short",day:"numeric",month:"short"})}</div>
        </div>
      </div>

      <div style={{display:"flex",borderBottom:"1px solid #2e2b22",background:"#141210"}}>
        {[["caja","💼 Caja"],["informes","📋 Informes"]].map(([k,label])=>(
          <button key={k} onClick={()=>setMainTab(k)}
            style={{flex:1,padding:"11px 0",background:"transparent",border:"none",borderBottom:mainTab===k?"2px solid #c9a84c":"2px solid transparent",color:mainTab===k?"#c9a84c":"#6a6047",fontSize:12,cursor:"pointer",letterSpacing:1}}>
            {label}
          </button>
        ))}
      </div>

      <div style={{maxWidth:820,margin:"0 auto",padding:"18px 12px"}}>
        {mainTab==="caja"&&<>
          <ResumenConsolidado/>
          <div style={{display:"grid",gridTemplateColumns:"1fr 1fr",gap:14}}>
            {LOCALES.map(local=><CajaLocal key={local.id} local={local}/>)}
          </div>
        </>}
        {mainTab==="informes"&&<Informes/>}
        <div style={{marginTop:12,fontSize:9,color:"#3a3520",textAlign:"center",letterSpacing:1}}>
          Toca el nombre para renombrarlo · Punto verde = día con datos · Punto rojo = descuadre
        </div>
      </div>
    </div>
  );
}

const inp={background:"#151310",border:"1px solid #2e2b22",borderRadius:7,padding:"9px 11px",color:"#e8e0cc",fontSize:12,outline:"none",fontFamily:"Georgia,serif"};
const btnPri={border:"none",color:"#0f0e0b",padding:"9px 16px",borderRadius:7,fontSize:12,fontWeight:"bold",cursor:"pointer",whiteSpace:"nowrap"};
const btnSec={background:"transparent",border:"1px solid #3a3520",color:"#8a7f5e",padding:"9px 12px",borderRadius:7,fontSize:12,cursor:"pointer",whiteSpace:"nowrap"};