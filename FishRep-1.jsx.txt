import { useState, useRef, useCallback } from "react";
import {
  LineChart, Line, BarChart, Bar, XAxis, YAxis, CartesianGrid,
  Tooltip, Legend, ResponsiveContainer, ComposedChart, Area
} from "recharts";
import * as XLSX from "xlsx";

/* ═══════════════════════════════════
   TEMPLATE GENERATOR
═══════════════════════════════════ */
function gerarPlanilhaModelo() {
  const wb = XLSX.utils.book_new();

  const addSheet = (name, rows) => {
    const ws = XLSX.utils.aoa_to_sheet(rows);
    XLSX.utils.book_append_sheet(wb, ws, name);
  };

  addSheet("T1_RGS_Order_Month", [
    ["Table 1. Monthly Relative Gonadosomatic Index (RGS) by Order"],
    ["Upper Paraguay River basin, northern Pantanal. Values: Mean +/- SD. n = sample size."],
    [],
    ["Month","Season","Siluriformes n","Siluriformes Mean RGS","Siluriformes SD","Siluriformes Median","Characiformes n","Characiformes Mean RGS","Characiformes SD","Characiformes Median","All Species n","All Species Mean RGS"],
    ["Oct","Pre-flood",143,0.973,0.816,0.641,216,1.575,2.24,0.672,297,1.339],
    ["Nov","Pre-flood",224,1.301,1.254,0.812,257,2.656,3.898,1.082,481,2.035],
    ["Dec","Flood rising",230,1.806,2.164,0.92,231,2.719,4.605,1.084,461,2.263],
    ["Jan","Flood peak",221,1.199,1.526,0.751,308,1.238,3.063,0.494,529,1.222],
    ["Feb","Flood peak",100,1.017,1.398,0.531,227,0.436,0.911,0.206,327,0.626],
    ["Mar","Flood peak",59,0.439,0.317,0.362,197,0.300,0.287,0.221,256,0.335],
    ["Apr","Recession",23,0.450,0.269,0.399,125,0.228,0.188,0.181,148,0.267],
    ["May","Dry season",38,0.347,0.172,0.311,104,0.244,0.184,0.210,142,0.271],
    ["Jun","Dry season",66,0.388,0.205,0.349,124,0.185,0.188,0.141,190,0.260],
    ["Jul","Dry season",47,0.299,0.227,0.248,128,0.244,0.230,0.193,175,0.264],
    ["Aug","Dry season",34,0.394,0.268,0.338,111,0.388,0.594,0.225,145,0.390],
    ["Sep","Pre-flood",27,0.436,0.239,0.386,117,0.712,1.114,0.365,206,0.639],
  ]);

  addSheet("T2_RGS_Species", [
    ["Table 2. Mean Relative Gonadosomatic Index (RGS) and body metrics by species"],
    ["Upper Paraguay River basin, northern Pantanal. Species sorted by descending mean RGS."],
    [],
    ["Species","Order","n total","n RGS","Mean RGS","SD RGS","Median RGS","Min RGS","Max RGS","Mean TL (cm)","SD TL","Mean TW (g)","n Female","n Male"],
    ["Prochilodus lineatus","Characiformes",153,153,3.222,5.594,1.210,0.031,22.181,34.8,3.8,452,83,70],
    ["Schizodon borellii","Characiformes",73,73,2.449,3.185,0.961,0.053,15.972,32.1,4.1,380,39,34],
    ["Ageneiosus brevifilis","Siluriformes",113,113,2.352,1.582,1.994,0.166,10.843,47.3,5.5,672,54,59],
    ["Pseudoplatystoma fasciatum","Siluriformes",226,226,1.570,1.921,0.955,0.012,14.222,83.8,9.0,5870,115,111],
    ["Leporinus friderici","Characiformes",15,15,1.032,1.336,0.551,0.095,5.014,31.8,3.2,310,8,7],
    ["Pinirampus pirinampu","Siluriformes",162,162,1.010,1.143,0.583,0.061,7.819,66.5,6.4,3220,80,82],
    ["Piaractus mesopotamicus","Characiformes",2145,2145,0.931,2.173,0.318,0.007,23.439,44.4,4.2,1410,1032,904],
    ["Pseudoplatystoma corruscans","Siluriformes",652,652,0.837,0.933,0.518,0.013,7.441,89.3,12.0,10890,330,322],
    ["Salminus brasiliensis","Characiformes",108,108,0.721,1.222,0.315,0.028,7.330,52.6,7.8,1980,57,51],
    ["Sorubim lima","Siluriformes",14,14,0.553,0.215,0.504,0.202,0.961,44.2,5.1,740,7,7],
    ["Hemisorubim platyrhynchus","Siluriformes",60,60,0.426,0.274,0.355,0.061,1.391,51.8,6.3,1470,29,31],
    ["Paulicea luetkeni","Siluriformes",48,48,0.309,0.196,0.272,0.053,0.888,78.4,11.2,7840,24,24],
  ]);

  addSheet("T3_RGS_Sex_Month", [
    ["Table 3. Monthly RGS by sex for key species"],
    ["Pseudoplatystoma corruscans and Piaractus mesopotamicus"],
    [],
    ["Species","Month","Season","Female n","Female Mean RGS","Female SD","Male n","Male Mean RGS","Male SD","U statistic","p-value"],
    ["Pseudoplatystoma corruscans","Oct","Pre-flood",70,0.891,0.880,73,0.847,0.708,"—","—"],
    ["Pseudoplatystoma corruscans","Nov","Pre-flood",115,1.012,1.211,109,0.952,0.931,"—","—"],
    ["Pseudoplatystoma corruscans","Dec","Flood rising",118,1.104,1.441,112,0.894,0.971,"—","—"],
    ["Pseudoplatystoma corruscans","Jan","Flood peak",108,0.812,1.022,99,0.721,0.887,"—","—"],
    ["Pseudoplatystoma corruscans","Feb","Flood peak",52,0.614,0.910,48,0.531,0.708,"—","—"],
    ["Pseudoplatystoma corruscans","Mar","Flood peak",31,0.412,0.291,28,0.371,0.301,"—","—"],
    ["Pseudoplatystoma corruscans","Apr","Recession",12,0.381,0.221,11,0.341,0.198,"—","—"],
    ["Pseudoplatystoma corruscans","May","Dry season",19,0.333,0.165,19,0.298,0.155,"—","—"],
    ["Pseudoplatystoma corruscans","Jun","Dry season",33,0.371,0.198,33,0.355,0.182,"—","—"],
    ["Pseudoplatystoma corruscans","Jul","Dry season",24,0.285,0.214,23,0.271,0.201,"—","—"],
    ["Pseudoplatystoma corruscans","Aug","Dry season",17,0.371,0.244,17,0.352,0.231,"—","—"],
    ["Pseudoplatystoma corruscans","Sep","Pre-flood",14,0.418,0.221,13,0.399,0.198,"—","—"],
    ["Piaractus mesopotamicus","Oct","Pre-flood",110,2.210,3.440,96,0.282,0.298,"—","—"],
    ["Piaractus mesopotamicus","Nov","Pre-flood",131,3.411,4.881,112,0.341,0.388,"—","—"],
    ["Piaractus mesopotamicus","Dec","Flood rising",118,3.788,5.912,99,0.351,0.401,"—","—"],
    ["Piaractus mesopotamicus","Jan","Flood peak",162,1.621,3.998,128,0.289,0.301,"—","—"],
    ["Piaractus mesopotamicus","Feb","Flood peak",112,0.481,0.921,105,0.221,0.188,"—","—"],
    ["Piaractus mesopotamicus","Mar","Flood peak",99,0.281,0.261,91,0.199,0.201,"—","—"],
    ["Piaractus mesopotamicus","Apr","Recession",62,0.221,0.181,57,0.181,0.148,"—","—"],
    ["Piaractus mesopotamicus","May","Dry season",51,0.244,0.188,47,0.201,0.162,"—","—"],
    ["Piaractus mesopotamicus","Jun","Dry season",62,0.178,0.172,59,0.162,0.153,"—","—"],
    ["Piaractus mesopotamicus","Jul","Dry season",62,0.233,0.218,60,0.208,0.191,"—","—"],
    ["Piaractus mesopotamicus","Aug","Dry season",54,0.362,0.481,51,0.238,0.221,"—","—"],
    ["Piaractus mesopotamicus","Sep","Pre-flood",57,0.688,1.021,55,0.381,0.442,"—","—"],
  ]);

  addSheet("T4_GonadStage_Freq", [
    ["Table 4. Monthly frequency (%) of gonadal maturation stages by order"],
    ["REP=Resting; ESV=Developing; MAT=Mature; EMA=Spawning"],
    [],
    ["Order","Month","Season","n total","REP %","ESV %","MAT %","EMA %"],
    ["Siluriformes","Oct","Pre-flood",143,22.4,14.7,0.7,62.3],
    ["Siluriformes","Nov","Pre-flood",224,15.2,10.3,1.0,73.5],
    ["Siluriformes","Dec","Flood rising",230,28.3,12.6,0.7,58.4],
    ["Siluriformes","Jan","Flood peak",221,50.2,13.1,1.5,35.2],
    ["Siluriformes","Feb","Flood peak",100,74.0,15.0,1.0,10.1],
    ["Siluriformes","Mar","Flood peak",59,98.3,1.7,0.0,0.0],
    ["Siluriformes","Apr","Recession",23,100.0,0.0,0.0,0.0],
    ["Siluriformes","May","Dry season",38,100.0,0.0,0.0,0.0],
    ["Siluriformes","Jun","Dry season",66,100.0,0.0,0.0,0.0],
    ["Siluriformes","Jul","Dry season",47,100.0,0.0,0.0,0.0],
    ["Siluriformes","Aug","Dry season",34,76.5,23.5,0.0,23.5],
    ["Siluriformes","Sep","Pre-flood",27,62.9,22.2,0.0,14.8],
    ["Characiformes","Oct","Pre-flood",216,18.1,4.6,0.0,77.2],
    ["Characiformes","Nov","Pre-flood",257,34.2,11.3,0.0,54.5],
    ["Characiformes","Dec","Flood rising",231,67.5,14.7,0.3,17.5],
    ["Characiformes","Jan","Flood peak",308,97.7,1.0,0.3,1.3],
    ["Characiformes","Feb","Flood peak",227,99.6,0.4,0.0,0.0],
    ["Characiformes","Mar","Flood peak",197,99.5,0.5,0.0,0.0],
    ["Characiformes","Apr","Recession",125,100.0,0.0,0.0,0.0],
    ["Characiformes","May","Dry season",104,100.0,0.0,0.0,0.0],
    ["Characiformes","Jun","Dry season",124,100.0,0.0,0.0,0.0],
    ["Characiformes","Jul","Dry season",128,99.2,0.8,0.0,0.0],
    ["Characiformes","Aug","Dry season",111,95.5,4.5,0.0,0.0],
    ["Characiformes","Sep","Pre-flood",117,52.1,4.3,0.0,43.6],
  ]);

  addSheet("T5_Environmental_Month", [
    ["Table 5. Monthly environmental variables"],
    ["Values: mean +/- SD for limnological variables"],
    [],
    ["Month","Season","n samples","Water table Mean (m)","Water table Max (m)","Water table Min (m)","Temperature Mean (C)","Temperature SD","DO Mean (mg/L)","pH Mean","Conductivity Mean (uS/cm)","Secchi disk Mean (cm)"],
    ["Oct","Pre-flood",297,1.47,1.76,1.24,29.8,1.2,6.9,7.31,68.2,26.4],
    ["Nov","Pre-flood",481,1.82,2.21,1.51,30.1,1.4,6.7,7.25,72.1,28.2],
    ["Dec","Flood rising",461,2.44,2.98,1.89,29.4,1.6,6.8,7.18,75.3,30.1],
    ["Jan","Flood peak",529,3.15,3.84,2.56,28.9,1.8,6.2,7.11,80.4,34.2],
    ["Feb","Flood peak",327,3.82,4.41,3.21,28.4,1.5,5.9,7.08,82.1,38.4],
    ["Mar","Flood peak",256,4.21,5.03,3.61,28.1,1.4,5.8,7.05,84.3,40.2],
    ["Apr","Recession",148,3.62,4.12,3.01,27.6,1.8,6.1,7.12,78.4,38.1],
    ["May","Dry season",142,2.81,3.21,2.41,26.2,2.1,6.4,7.18,69.2,33.4],
    ["Jun","Dry season",190,1.98,2.31,1.71,24.4,2.4,6.7,7.22,63.4,29.8],
    ["Jul","Dry season",175,1.44,1.71,1.21,22.8,2.8,6.9,7.31,61.2,27.4],
    ["Aug","Dry season",145,1.12,1.38,0.92,24.6,2.2,6.8,7.28,62.1,25.8],
    ["Sep","Pre-flood",206,1.21,1.48,0.98,27.8,1.6,6.7,7.24,63.8,24.2],
  ]);

  addSheet("T6_Statistical_Tests", [
    ["Table 6. Summary of statistical tests"],
    ["All non-parametric tests. Significance level: alpha = 0.05."],
    [],
    ["Test","Variable","Group 1","Group 2","Statistic","df / n","p-value","Interpretation"],
    ["Kruskal-Wallis","RGS across months","Siluriformes","12 monthly groups","H = 226.21","df = 11; n = 1212","p < 0.0001","Strongly seasonal reproduction"],
    ["Kruskal-Wallis","RGS across months","Characiformes","12 monthly groups","H = 481.21","df = 11; n = 2145","p < 0.0001","Strongly seasonal reproduction"],
    ["Mann-Whitney U","RGS Female vs Male","Siluriformes Female (mean=1.168)","Siluriformes Male (mean=1.017)","U = 202370","nF=627; nM=585","p = 0.0018","Females invest more gonadal mass"],
    ["Mann-Whitney U","RGS Female vs Male","Characiformes Female (mean=1.826)","Characiformes Male (mean=0.308)","U = 862224","nF=1162; nM=983","p < 0.0001","Strong sexual dimorphism"],
    ["Mann-Whitney U","RGS Female > Male","P. mesopotamicus Female (1.482)","P. mesopotamicus Male (0.303)","U = 688834","nF=1032; nM=904","p < 0.0001","5x higher female RGS"],
    ["Mann-Whitney U","RGS Female vs Male","P. corruscans Female (0.869)","P. corruscans Male (0.805)","U = 50866","nF=330; nM=322","p = 0.347 (NS)","No significant sexual dimorphism"],
    ["Spearman rs","RGS vs Water Temp","RGS (all species)","Water temperature (C)","rs = 0.244","n = 1391","p < 0.0001","Primary environmental driver"],
    ["Spearman rs","RGS vs Water Table","RGS (all species)","River water table (m)","rs = -0.096","n = 1391","p = 0.0004","Negative: activation precedes flood peak"],
    ["Spearman rs","RGS vs pH","RGS (all species)","pH","rs = -0.014","n = 1391","p = 0.609 (NS)","No significant correlation"],
    ["Spearman rs","RGS vs DO","RGS (all species)","Dissolved oxygen (mg/L)","rs = 0.038","n = 1391","p = 0.162 (NS)","No significant correlation"],
  ]);

  addSheet("T7_Sample_Overview", [
    ["Table 7. Sample overview by species, order and sex"],
    ["Upper Paraguay River basin, northern Pantanal."],
    [],
    ["Species","Order","Total n","Female n","Male n","Mean TL cm","SD TL","Min TL","Max TL","Mean TW g","Mean RGS","SD RGS","Reproductive guild"],
    ["Piaractus mesopotamicus","Characiformes",2145,1032,904,44.4,4.2,36,67.5,1410,0.931,2.173,"Periodic - migratory"],
    ["Pseudoplatystoma corruscans","Siluriformes",652,330,322,89.3,12.0,67,130,10890,0.837,0.933,"Periodic - migratory"],
    ["Pseudoplatystoma fasciatum","Siluriformes",226,115,111,83.8,9.0,62,105,5870,1.570,1.921,"Periodic - migratory"],
    ["Pinirampus pirinampu","Siluriformes",162,80,82,66.5,6.4,48,92.3,3220,1.010,1.143,"Periodic - migratory"],
    ["Prochilodus lineatus","Characiformes",153,83,70,34.8,3.8,28,49,452,3.222,5.594,"Periodic - migratory"],
    ["Ageneiosus brevifilis","Siluriformes",113,54,59,47.3,5.5,35.8,67.5,672,2.352,1.582,"Opportunistic"],
    ["Salminus brasiliensis","Characiformes",108,57,51,52.6,7.8,38,78,1980,0.721,1.222,"Periodic - migratory"],
    ["Schizodon borellii","Characiformes",73,39,34,32.1,4.1,24,48.5,380,2.449,3.185,"Periodic"],
    ["Hemisorubim platyrhynchus","Siluriformes",60,29,31,51.8,6.3,38.2,71.4,1470,0.426,0.274,"Periodic"],
    ["Paulicea luetkeni","Siluriformes",48,24,24,78.4,11.2,58,108,7840,0.309,0.196,"Periodic - migratory"],
    ["Leporinus friderici","Characiformes",15,8,7,31.8,3.2,26,41,310,1.032,1.336,"Periodic"],
    ["Sorubim lima","Siluriformes",14,7,7,44.2,5.1,34,54,740,0.553,0.215,"Periodic"],
    ["TOTAL","—",3357,1789,1568,"—","—","—","—","—","—","—","—"],
  ]);

  XLSX.writeFile(wb, "Modelo_FishRep.xlsx");
}

/* ═══════════════════════════════════
   PARSE XLSX
═══════════════════════════════════ */
const MN = ["Out","Nov","Dez","Jan","Fev","Mar","Abr","Mai","Jun","Jul","Ago","Set"];
const MAP_SZN = {Oct:"Pré-cheia",Nov:"Pré-cheia",Dec:"Cheia",Jan:"Cheia plena",Feb:"Cheia plena",Mar:"Cheia plena",Apr:"Vazante",May:"Seca",Jun:"Seca",Jul:"Seca",Aug:"Seca",Sep:"Pré-cheia"};

async function parseFile(file) {
  const buf = await file.arrayBuffer();
  const wb = XLSX.read(buf, { type: "array" });
  const rows = name => {
    const ws = wb.Sheets[name]; if (!ws) return null;
    return XLSX.utils.sheet_to_json(ws, { header: 1, defval: null }).slice(3).filter(r => r.some(c => c != null));
  };
  const r1=rows("T1_RGS_Order_Month"), r2=rows("T2_RGS_Species"), r3=rows("T3_RGS_Sex_Month"),
        r4=rows("T4_GonadStage_Freq"), r5=rows("T5_Environmental_Month"), r6=rows("T6_Statistical_Tests"), r7=rows("T7_Sample_Overview");
  return {
    t1: r1?.map((r,i)=>({m:MN[i]??r[0],s:MAP_SZN[r[1]]??r[1],sil:+r[3]||0,cha:+r[7]||0,all:+r[11]||0,silN:+r[2]||0,chaN:+r[6]||0})).slice(0,12)??[],
    t2: r2?.map(r=>({sp:r[0],sn:(r[0]||"").split(" ").slice(0,2).map((w,i)=>i===0?w[0]+".":w).join(" "),ord:r[1],n:+r[2]||0,rgs:+r[4]||0,sd:+r[5]||0,medRGS:+r[6]||0,tl:+r[9]||0,tw:+r[11]||0,nF:+r[12]||0,nM:+r[13]||0})).filter(d=>typeof d.sp==="string"&&d.sp.length>3)??[],
    t3c: r3?.filter(r=>r[0]==="Pseudoplatystoma corruscans").map((r,i)=>({m:MN[i]??r[1],fRGS:+r[4]||0,mRGS:+r[7]||0}))??[],
    t3p: r3?.filter(r=>r[0]==="Piaractus mesopotamicus").map((r,i)=>({m:MN[i]??r[1],fRGS:+r[4]||0,mRGS:+r[7]||0}))??[],
    t4s: r4?.filter(r=>r[0]==="Siluriformes").map((r,i)=>({m:MN[i]??r[1],n:+r[3]||0,REP:+r[4]||0,ESV:+r[5]||0,MAT:+r[6]||0,EMA:+r[7]||0}))??[],
    t4c: r4?.filter(r=>r[0]==="Characiformes").map((r,i)=>({m:MN[i]??r[1],n:+r[3]||0,REP:+r[4]||0,ESV:+r[5]||0,MAT:+r[6]||0,EMA:+r[7]||0}))??[],
    t5: r5?.map((r,i)=>({m:MN[i]??r[0],wt:+r[3]||0,temp:+r[6]||0,do:+r[8]||0,ph:+r[9]||0,cond:+r[10]||0,sec:+r[11]||0})).slice(0,12)??[],
    t6: r6?.map(r=>({teste:r[0],variavel:r[1],grupo:r[2],stat:r[4],n:r[5],p:r[6],sig:!String(r[6]).includes("NS"),interp:r[7]})).filter(d=>d.teste)??[],
    t7: r7?.map(r=>({sp:r[0],ord:r[1],n:+r[2]||0,nF:+r[3]||0,nM:+r[4]||0,tl:+r[5]||0,tw:+r[9]||0,rgs:+r[10]||0,sd:+r[11]||0,guild:r[12]})).filter(d=>typeof d.sp==="string"&&d.sp!=="TOTAL"&&+d.n>0)??[],
  };
}

/* ═══════════════════════════════════
   COLORS & UTILS
═══════════════════════════════════ */
const SIL="#e07b39", CHA="#2d9e5f", FEM="#d63384", MAL="#0ea5e9";
const SEASON_COLORS={"Pré-cheia":"#fb923c","Cheia":"#3b82f6","Cheia plena":"#1d4ed8","Vazante":"#a855f7","Seca":"#94a3b8"};
const f2=v=>v!=null&&!isNaN(+v)?Number(v).toFixed(2):"–";
const f3=v=>v!=null&&!isNaN(+v)?Number(v).toFixed(3):"–";

const dlCSV=(rows,fname)=>{
  if(!rows?.length)return;
  const keys=Object.keys(rows[0]);
  const csv=[keys.join(","),...rows.map(r=>keys.map(k=>JSON.stringify(r[k]??"")).join(","))].join("\n");
  const a=Object.assign(document.createElement("a"),{href:"data:text/csv;charset=utf-8,"+encodeURIComponent(csv),download:fname+".csv"});
  document.body.appendChild(a);a.click();document.body.removeChild(a);
};
const dlSVG=(ref,fname)=>{
  const svg=ref.current?.querySelector("svg");if(!svg)return;
  const cl=svg.cloneNode(true);cl.setAttribute("xmlns","http://www.w3.org/2000/svg");
  const st=document.createElementNS("http://www.w3.org/2000/svg","style");
  st.textContent="text{font-family:Trebuchet MS,sans-serif}";
  cl.insertBefore(st,cl.firstChild);
  const blob=new Blob([new XMLSerializer().serializeToString(cl)],{type:"image/svg+xml"});
  const a=Object.assign(document.createElement("a"),{href:URL.createObjectURL(blob),download:fname+".svg"});
  document.body.appendChild(a);a.click();document.body.removeChild(a);
};

/* ─────── small UI ─────── */
const Tip=({active,payload,label})=>{
  if(!active||!payload?.length)return null;
  return <div style={{background:"#fff",border:"1.5px solid #bfdbfe",borderRadius:10,padding:"8px 14px",fontSize:12,boxShadow:"0 4px 18px rgba(0,80,160,.12)"}}>
    <div style={{fontWeight:800,color:"#1e3a5f",marginBottom:5}}>{label}</div>
    {payload.map((p,i)=><div key={i} style={{display:"flex",gap:8,alignItems:"center",color:p.color,marginBottom:2}}>
      <span style={{width:8,height:8,borderRadius:"50%",background:p.color,display:"inline-block",flexShrink:0}}/>
      {p.name}: <strong>{typeof p.value==="number"?p.value.toFixed(3):p.value}</strong>
    </div>)}
  </div>;
};
const DLRow=({svgRef,svgName,csvRows,csvName})=>(
  <div style={{display:"flex",gap:6}}>
    {svgRef&&<button onClick={()=>dlSVG(svgRef,svgName)} style={BS("#0ea5e9")}>⬇ SVG</button>}
    {csvRows&&<button onClick={()=>dlCSV(csvRows,csvName)} style={BS("#2d9e5f")}>⬇ CSV</button>}
  </div>
);
const BS=c=>({background:c,color:"#fff",border:"none",borderRadius:7,padding:"5px 12px",fontSize:11,fontWeight:700,cursor:"pointer"});

/* shared styles */
const card={background:"#fff",borderRadius:14,boxShadow:"0 2px 16px rgba(0,80,160,.07)",padding:"20px",marginBottom:16};
const tbl={width:"100%",borderCollapse:"collapse",fontSize:12};
const TH={background:"#f1f7fd",padding:"8px 11px",textAlign:"left",fontWeight:800,color:"#1e3a5f",borderBottom:"2px solid #dbeafe",whiteSpace:"nowrap",fontSize:10.5,textTransform:"uppercase",letterSpacing:".04em"};
const TD={padding:"7px 11px",borderBottom:"1px solid #f1f5f9",color:"#374151"};
const BSil={padding:"2px 8px",borderRadius:20,fontSize:10,fontWeight:700,background:"#fff0e5",color:"#c05a17"};
const BCha={padding:"2px 8px",borderRadius:20,fontSize:10,fontWeight:700,background:"#e8f7ee",color:"#1a6b3d"};
const SEL={border:"1.5px solid #bfdbfe",borderRadius:8,padding:"5px 10px",fontSize:12,fontWeight:700,color:"#1e3a5f",background:"#fff",cursor:"pointer",outline:"none"};

/* ═══════════════════════════════════
   TAB SECTIONS
═══════════════════════════════════ */
function TabOverview({d}){
  const tN=d.t7.reduce((s,r)=>s+(r.n||0),0);
  const tF=d.t7.reduce((s,r)=>s+(r.nF||0),0);
  const tM=d.t7.reduce((s,r)=>s+(r.nM||0),0);
  const cards=[
    {ic:"🐟",v:tN.toLocaleString("pt-BR"),l:"Espécimes totais",c:"#0077b6"},
    {ic:"🧬",v:d.t7.length,l:"Espécies",c:"#2d9e5f"},
    {ic:"♀",v:tF.toLocaleString("pt-BR"),l:"Fêmeas",c:"#d63384"},
    {ic:"♂",v:tM.toLocaleString("pt-BR"),l:"Machos",c:"#0ea5e9"},
  ];
  return(
    <div>
      <div style={{display:"flex",gap:12,flexWrap:"wrap",marginBottom:18}}>
        {cards.map((c,i)=>(
          <div key={i} style={{background:"#fff",borderRadius:12,padding:"14px 18px",textAlign:"center",boxShadow:"0 2px 10px rgba(0,80,160,.07)",flex:1,minWidth:100,maxWidth:160,borderTop:`3px solid ${c.c}`}}>
            <div style={{fontSize:24,marginBottom:4}}>{c.ic}</div>
            <div style={{fontSize:20,fontWeight:900,color:c.c,lineHeight:1}}>{c.v}</div>
            <div style={{fontSize:10,color:"#64748b",marginTop:4,fontWeight:600}}>{c.l}</div>
          </div>
        ))}
      </div>
      <div style={card}>
        <div style={{fontWeight:800,color:"#1e3a5f",fontSize:14,marginBottom:12}}>Panorama por Especie</div>
        <div style={{overflowX:"auto"}}>
          <table style={tbl}>
            <thead><tr>{["Especie","Ordem","n","n ♀","n ♂","CT (cm)","PM (g)","IGS","DP IGS","Guilda"].map(h=><th key={h} style={TH}>{h}</th>)}</tr></thead>
            <tbody>{[...d.t7].sort((a,b)=>b.n-a.n).map((r,i)=>(
              <tr key={i} style={{background:i%2===0?"#f8fbff":"#fff"}}>
                <td style={{...TD,fontStyle:"italic",fontWeight:700,color:"#1e3a5f"}}>{r.sp}</td>
                <td style={TD}><span style={r.ord==="Siluriformes"?BSil:BCha}>{r.ord==="Siluriformes"?"Silu.":"Char."}</span></td>
                <td style={TD}>{r.n}</td><td style={TD}>{r.nF}</td><td style={TD}>{r.nM}</td>
                <td style={TD}>{f2(r.tl)}</td><td style={TD}>{r.tw?.toLocaleString("pt-BR")??"-"}</td>
                <td style={{...TD,fontWeight:700,color:r.rgs>2?"#e03c3c":r.rgs>1?"#e07b39":"#374151"}}>{f3(r.rgs)}</td>
                <td style={TD}>{f3(r.sd)}</td><td style={TD}>{r.guild}</td>
              </tr>
            ))}</tbody>
          </table>
        </div>
        <div style={{marginTop:8}}><button onClick={()=>dlCSV(d.t7,"panorama_especies")} style={BS("#2d9e5f")}>⬇ CSV</button></div>
      </div>
    </div>
  );
}

function TabRGSMensal({d}){
  const ref=useRef();
  return(
    <div>
      <div style={card}>
        <div style={{display:"flex",justifyContent:"space-between",flexWrap:"wrap",gap:10,marginBottom:12}}>
          <div><div style={{fontWeight:800,color:"#1e3a5f",fontSize:14}}>IGS Relativo por Mes e Ordem</div>
          <div style={{fontSize:11,color:"#64748b"}}>Pantanal Norte · 12 meses hidrológicos</div></div>
          <DLRow svgRef={ref} svgName="IGS_mensal" csvRows={d.t1} csvName="IGS_mensal"/>
        </div>
        <div ref={ref} style={{width:"100%",height:320}}>
          <ResponsiveContainer><ComposedChart data={d.t1} margin={{top:5,right:20,bottom:5,left:0}}>
            <CartesianGrid strokeDasharray="3 3" stroke="#e2e8f0"/>
            <XAxis dataKey="m" tick={{fontSize:12,fontWeight:600}}/>
            <YAxis tick={{fontSize:11}} label={{value:"IGS Medio",angle:-90,position:"insideLeft",style:{fontSize:11,fill:"#64748b"}}}/>
            <Tooltip content={<Tip/>}/><Legend wrapperStyle={{fontSize:12}}/>
            <Area type="monotone" dataKey="all" name="Geral" fill="#dbeafe" stroke="#93c5fd" strokeWidth={1} dot={false} fillOpacity={0.5}/>
            <Line type="monotone" dataKey="sil" name="Siluriformes" stroke={SIL} strokeWidth={2.5} dot={{r:4,fill:SIL}} activeDot={{r:6}}/>
            <Line type="monotone" dataKey="cha" name="Characiformes" stroke={CHA} strokeWidth={2.5} dot={{r:4,fill:CHA}} activeDot={{r:6}}/>
          </ComposedChart></ResponsiveContainer>
        </div>
        <div style={{display:"flex",gap:8,flexWrap:"wrap",marginTop:6}}>
          {Object.entries(SEASON_COLORS).map(([s,c])=>(
            <span key={s} style={{display:"flex",alignItems:"center",gap:4,fontSize:10,color:"#374151"}}>
              <span style={{width:10,height:10,borderRadius:3,background:c,opacity:.7,display:"inline-block"}}/>{s}
            </span>
          ))}
        </div>
      </div>
      <div style={card}>
        <div style={{fontWeight:800,color:"#1e3a5f",fontSize:14,marginBottom:10}}>Tabela Mensal</div>
        <div style={{overflowX:"auto"}}>
          <table style={tbl}>
            <thead><tr>{["Mes","Estacao","n Silu","IGS Silu","n Char","IGS Char","n Total","IGS Total"].map(h=><th key={h} style={TH}>{h}</th>)}</tr></thead>
            <tbody>{d.t1.map((r,i)=>(
              <tr key={i} style={{background:i%2===0?"#f8fbff":"#fff"}}>
                <td style={{...TD,fontWeight:700}}>{r.m}</td>
                <td style={TD}><span style={{padding:"2px 8px",borderRadius:20,fontSize:10,fontWeight:700,background:(SEASON_COLORS[r.s]||"#94a3b8")+"22",color:SEASON_COLORS[r.s]||"#64748b"}}>{r.s}</span></td>
                <td style={TD}>{r.silN}</td><td style={{...TD,color:SIL,fontWeight:700}}>{f3(r.sil)}</td>
                <td style={TD}>{r.chaN}</td><td style={{...TD,color:CHA,fontWeight:700}}>{f3(r.cha)}</td>
                <td style={TD}>{r.silN+r.chaN}</td><td style={{...TD,fontWeight:700}}>{f3(r.all)}</td>
              </tr>
            ))}</tbody>
          </table>
        </div>
      </div>
    </div>
  );
}

function TabRGSEspecies({d}){
  const ref=useRef();
  const sorted=[...d.t2].sort((a,b)=>b.rgs-a.rgs);
  return(
    <div>
      <div style={card}>
        <div style={{display:"flex",justifyContent:"space-between",flexWrap:"wrap",gap:10,marginBottom:12}}>
          <div><div style={{fontWeight:800,color:"#1e3a5f",fontSize:14}}>IGS Medio por Especie</div>
          <div style={{fontSize:11,color:"#64748b"}}>Ordenado por valor decrescente</div></div>
          <DLRow svgRef={ref} svgName="IGS_especies" csvRows={d.t2} csvName="IGS_especies"/>
        </div>
        <div ref={ref} style={{width:"100%",height:Math.max(260,sorted.length*32)}}>
          <ResponsiveContainer><BarChart data={sorted} layout="vertical" margin={{top:5,right:30,bottom:5,left:130}}>
            <CartesianGrid strokeDasharray="3 3" stroke="#e2e8f0"/>
            <XAxis type="number" tick={{fontSize:11}}/>
            <YAxis type="category" dataKey="sn" tick={{fontSize:11,fontStyle:"italic"}} width={125}/>
            <Tooltip content={<Tip/>}/>
            <Bar dataKey="rgs" name="IGS Medio" radius={[0,5,5,0]}>
              {sorted.map((r,i)=><rect key={i} fill={r.ord==="Siluriformes"?SIL:CHA}/>)}
            </Bar>
          </BarChart></ResponsiveContainer>
        </div>
      </div>
      <div style={card}>
        <div style={{fontWeight:800,color:"#1e3a5f",fontSize:14,marginBottom:10}}>Tabela de Metricas</div>
        <div style={{overflowX:"auto"}}>
          <table style={tbl}>
            <thead><tr>{["Especie","Ordem","n","IGS","DP","Mediana","CT (cm)","PM (g)","n ♀","n ♂"].map(h=><th key={h} style={TH}>{h}</th>)}</tr></thead>
            <tbody>{sorted.map((r,i)=>(
              <tr key={i} style={{background:i%2===0?"#f8fbff":"#fff"}}>
                <td style={{...TD,fontStyle:"italic",fontWeight:700,color:"#1e3a5f"}}>{r.sp}</td>
                <td style={TD}><span style={r.ord==="Siluriformes"?BSil:BCha}>{r.ord}</span></td>
                <td style={TD}>{r.n}</td>
                <td style={{...TD,fontWeight:700,color:r.rgs>2?"#e03c3c":r.rgs>1?"#e07b39":"#374151"}}>{f3(r.rgs)}</td>
                <td style={TD}>{f3(r.sd)}</td><td style={TD}>{f3(r.medRGS)}</td>
                <td style={TD}>{f2(r.tl)}</td><td style={TD}>{r.tw?.toLocaleString("pt-BR")??"-"}</td>
                <td style={TD}>{r.nF}</td><td style={TD}>{r.nM}</td>
              </tr>
            ))}</tbody>
          </table>
        </div>
      </div>
    </div>
  );
}

function TabSexo({d}){
  const [sp,setSp]=useState("corruscans");
  const data=sp==="corruscans"?d.t3c:d.t3p;
  const ref=useRef();
  return(
    <div>
      <div style={{display:"flex",gap:10,marginBottom:14,alignItems:"center"}}>
        <span style={{fontSize:12,fontWeight:700,color:"#1e3a5f"}}>Especie:</span>
        <select value={sp} onChange={e=>setSp(e.target.value)} style={SEL}>
          <option value="corruscans">Pseudoplatystoma corruscans (Siluriformes)</option>
          <option value="mesopotamicus">Piaractus mesopotamicus (Characiformes)</option>
        </select>
      </div>
      <div style={card}>
        <div style={{display:"flex",justifyContent:"space-between",flexWrap:"wrap",gap:10,marginBottom:12}}>
          <div><div style={{fontWeight:800,color:"#1e3a5f",fontSize:14,fontStyle:"italic"}}>{sp==="corruscans"?"Pseudoplatystoma corruscans":"Piaractus mesopotamicus"}</div>
          <div style={{fontSize:11,color:"#64748b"}}>IGS mensal por sexo</div></div>
          <DLRow svgRef={ref} svgName={`IGS_sexo_${sp}`} csvRows={data} csvName={`IGS_sexo_${sp}`}/>
        </div>
        <div ref={ref} style={{width:"100%",height:300}}>
          <ResponsiveContainer><LineChart data={data} margin={{top:5,right:20,bottom:5,left:0}}>
            <CartesianGrid strokeDasharray="3 3" stroke="#e2e8f0"/>
            <XAxis dataKey="m" tick={{fontSize:12,fontWeight:600}}/>
            <YAxis tick={{fontSize:11}} label={{value:"IGS Medio",angle:-90,position:"insideLeft",style:{fontSize:11,fill:"#64748b"}}}/>
            <Tooltip content={<Tip/>}/><Legend wrapperStyle={{fontSize:12}}/>
            <Line type="monotone" dataKey="fRGS" name="♀ Femeas" stroke={FEM} strokeWidth={2.5} dot={{r:4,fill:FEM}} activeDot={{r:6}}/>
            <Line type="monotone" dataKey="mRGS" name="♂ Machos" stroke={MAL} strokeWidth={2.5} dot={{r:4,fill:MAL}} activeDot={{r:6}} strokeDasharray="5 3"/>
          </LineChart></ResponsiveContainer>
        </div>
        <div style={{background:"#f8fbff",borderRadius:8,padding:"10px 14px",marginTop:10,fontSize:11.5,color:"#374151",borderLeft:`3px solid ${sp==="mesopotamicus"?"#d63384":"#94a3b8"}`}}>
          {sp==="mesopotamicus"?"Mann-Whitney U = 688.834; p < 0.0001 — IGS feminino ~5x maior que masculino.":"Mann-Whitney U = 50.866; p = 0.347 (NS) — Sem dimorfismo sexual significativo."}
        </div>
      </div>
    </div>
  );
}

function TabGonadas({d}){
  const [ord,setOrd]=useState("siluriformes");
  const data=ord==="siluriformes"?d.t4s:d.t4c;
  const ref=useRef();
  return(
    <div>
      <div style={{display:"flex",gap:10,marginBottom:14,alignItems:"center"}}>
        <span style={{fontSize:12,fontWeight:700,color:"#1e3a5f"}}>Ordem:</span>
        <select value={ord} onChange={e=>setOrd(e.target.value)} style={SEL}>
          <option value="siluriformes">Siluriformes</option>
          <option value="characiformes">Characiformes</option>
        </select>
      </div>
      <div style={card}>
        <div style={{display:"flex",justifyContent:"space-between",flexWrap:"wrap",gap:10,marginBottom:12}}>
          <div><div style={{fontWeight:800,color:"#1e3a5f",fontSize:14}}>Estadios Gonadais — {ord==="siluriformes"?"Siluriformes":"Characiformes"}</div>
          <div style={{fontSize:11,color:"#64748b"}}>REP=Repouso · ESV=Desenvolvimento · MAT=Maduro · EMA=Espermiacao</div></div>
          <DLRow svgRef={ref} svgName={`estadios_${ord}`} csvRows={data} csvName={`estadios_${ord}`}/>
        </div>
        <div ref={ref} style={{width:"100%",height:320}}>
          <ResponsiveContainer><BarChart data={data} margin={{top:5,right:20,bottom:5,left:0}}>
            <CartesianGrid strokeDasharray="3 3" stroke="#e2e8f0"/>
            <XAxis dataKey="m" tick={{fontSize:12,fontWeight:600}}/>
            <YAxis tick={{fontSize:11}} label={{value:"%",angle:-90,position:"insideLeft",style:{fontSize:11,fill:"#64748b"}}}/>
            <Tooltip content={<Tip/>}/><Legend wrapperStyle={{fontSize:12}}/>
            <Bar dataKey="REP" name="Repouso" stackId="a" fill="#94a3b8"/>
            <Bar dataKey="ESV" name="Desenvolvimento" stackId="a" fill="#38bdf8"/>
            <Bar dataKey="MAT" name="Maduro" stackId="a" fill="#fb923c"/>
            <Bar dataKey="EMA" name="Espermiacao" stackId="a" fill="#ef4444" radius={[4,4,0,0]}/>
          </BarChart></ResponsiveContainer>
        </div>
      </div>
    </div>
  );
}

function TabAmbiental({d}){
  const [vari,setVari]=useState("temp");
  const vars={temp:{label:"Temperatura (°C)",color:"#ef4444"},wt:{label:"Nivel do rio (m)",color:"#3b82f6"},do:{label:"OD (mg/L)",color:"#10b981"},cond:{label:"Condutividade (µS/cm)",color:"#8b5cf6"},ph:{label:"pH",color:"#f59e0b"},sec:{label:"Secchi (cm)",color:"#6b7280"}};
  const ref=useRef();
  return(
    <div>
      <div style={{display:"flex",gap:10,marginBottom:14,alignItems:"center",flexWrap:"wrap"}}>
        <span style={{fontSize:12,fontWeight:700,color:"#1e3a5f"}}>Variavel:</span>
        <select value={vari} onChange={e=>setVari(e.target.value)} style={SEL}>
          {Object.entries(vars).map(([k,v])=><option key={k} value={k}>{v.label}</option>)}
        </select>
      </div>
      <div style={card}>
        <div style={{display:"flex",justifyContent:"space-between",flexWrap:"wrap",gap:10,marginBottom:12}}>
          <div><div style={{fontWeight:800,color:"#1e3a5f",fontSize:14}}>{vars[vari].label} — Variacao Mensal</div></div>
          <DLRow svgRef={ref} svgName={`ambiental_${vari}`} csvRows={d.t5} csvName="variaveis_ambientais"/>
        </div>
        <div ref={ref} style={{width:"100%",height:280}}>
          <ResponsiveContainer><ComposedChart data={d.t5} margin={{top:5,right:20,bottom:5,left:0}}>
            <CartesianGrid strokeDasharray="3 3" stroke="#e2e8f0"/>
            <XAxis dataKey="m" tick={{fontSize:12,fontWeight:600}}/>
            <YAxis tick={{fontSize:11}}/>
            <Tooltip content={<Tip/>}/>
            <Area type="monotone" dataKey={vari} name={vars[vari].label} fill={vars[vari].color+"22"} stroke={vars[vari].color} strokeWidth={2.5} dot={{r:4,fill:vars[vari].color}} activeDot={{r:6}}/>
          </ComposedChart></ResponsiveContainer>
        </div>
      </div>
    </div>
  );
}

function TabEstatistica({d}){
  return(
    <div>
      <div style={card}>
        <div style={{fontWeight:800,color:"#1e3a5f",fontSize:14,marginBottom:12}}>Testes Estatisticos</div>
        <div style={{overflowX:"auto"}}>
          <table style={tbl}>
            <thead><tr>{["Teste","Variavel","Grupo","Estatistica","n / df","Valor p","Sig.","Interpretacao"].map(h=><th key={h} style={TH}>{h}</th>)}</tr></thead>
            <tbody>{d.t6.map((r,i)=>(
              <tr key={i} style={{background:i%2===0?"#f8fbff":"#fff"}}>
                <td style={{...TD,fontWeight:700,color:"#1e3a5f"}}>{r.teste}</td>
                <td style={TD}>{r.variavel}</td>
                <td style={{...TD,fontSize:11}}>{r.grupo}</td>
                <td style={{...TD,fontFamily:"monospace",fontSize:11,fontWeight:700}}>{r.stat}</td>
                <td style={{...TD,fontSize:11}}>{r.n}</td>
                <td style={{...TD,fontFamily:"monospace",fontWeight:700,color:r.sig?"#b91c1c":"#374151"}}>{r.p}</td>
                <td style={TD}><span style={r.sig?{background:"#fee2e2",color:"#b91c1c",padding:"2px 8px",borderRadius:20,fontSize:10,fontWeight:700}:{background:"#f3f4f6",color:"#6b7280",padding:"2px 8px",borderRadius:20,fontSize:10,fontWeight:700}}>{r.sig?"Sim ✓":"NS"}</span></td>
                <td style={{...TD,fontSize:11}}>{r.interp}</td>
              </tr>
            ))}</tbody>
          </table>
        </div>
        <div style={{marginTop:8}}><button onClick={()=>dlCSV(d.t6,"testes_estatisticos")} style={BS("#2d9e5f")}>⬇ CSV</button></div>
      </div>
    </div>
  );
}

/* ═══════════════════════════════════
   MAIN APP
═══════════════════════════════════ */
export default function FishRep(){
  const [data,setData]=useState(null);
  const [loading,setLoading]=useState(false);
  const [tab,setTab]=useState(0);
  const [drag,setDrag]=useState(false);
  const fileRef=useRef();

  const handleFile=useCallback(async(file)=>{
    if(!file||!file.name.endsWith(".xlsx"))return;
    setLoading(true);
    try{
      const d=await parseFile(file);
      setData(d);
      setTab(0);
    }catch(e){alert("Erro ao ler arquivo. Verifique se e o formato modelo correto.");}
    setLoading(false);
  },[]);

  const TABS=[
    {icon:"🏠",label:"Visao Geral"},
    {icon:"📈",label:"IGS Mensal"},
    {icon:"🐟",label:"Por Especie"},
    {icon:"⚥",label:"Sexo"},
    {icon:"🔬",label:"Estadios Gonadais"},
    {icon:"🌊",label:"Ambiental"},
    {icon:"📊",label:"Estatistica"},
  ];

  return(
    <div style={{minHeight:"100vh",background:"linear-gradient(135deg,#e0f2fe 0%,#f0fdf4 100%)",fontFamily:"'Trebuchet MS','Segoe UI',sans-serif",position:"relative",overflowX:"hidden"}}>
      {/* FISH ANIMATION */}
      {[
        {e:"🐟",top:"15%",d:"0s",  dur:"9s", dir:"r",sz:30},
        {e:"🐠",top:"60%",d:"4s",  dur:"13s",dir:"r",sz:24},
        {e:"🐡",top:"38%",d:"7s",  dur:"11s",dir:"l",sz:26},
        {e:"🐟",top:"78%",d:"2s",  dur:"15s",dir:"r",sz:18},
        {e:"🐠",top:"48%",d:"10s", dur:"12s",dir:"l",sz:21},
      ].map((f,i)=>(
        <div key={i} style={{position:"fixed",top:f.top,zIndex:0,pointerEvents:"none",userSelect:"none",fontSize:f.sz,animation:`${f.dir==="r"?"fishR":"fishL"} ${f.dur} ${f.d} linear infinite`,opacity:.6}}>{f.e}</div>
      ))}
      <style>{`
        @keyframes fishR{0%{transform:translateX(-80px)}100%{transform:translateX(110vw)}}
        @keyframes fishL{0%{transform:translateX(110vw) scaleX(-1)}100%{transform:translateX(-80px) scaleX(-1)}}
        @keyframes fadeUp{from{opacity:0;transform:translateY(14px)}to{opacity:1;transform:translateY(0)}}
        .fu{animation:fadeUp .4s ease forwards}
      `}</style>

      {/* HEADER */}
      <div style={{background:"linear-gradient(135deg,#0077b6 0%,#00b4d8 60%,#2d9e5f 100%)",padding:"18px 24px",position:"relative",zIndex:10,boxShadow:"0 4px 20px rgba(0,80,160,.25)"}}>
        <div style={{maxWidth:1100,margin:"0 auto",display:"flex",alignItems:"center",justifyContent:"space-between",gap:16,flexWrap:"wrap"}}>
          <div style={{display:"flex",alignItems:"center",gap:12}}>
            <div style={{fontSize:42,lineHeight:1}}>🐟</div>
            <div>
              <div style={{color:"#fff",fontSize:25,fontWeight:900,letterSpacing:"-.5px",lineHeight:1}}>
                Fish<span style={{color:"#90e0ef"}}>Rep</span> <span style={{fontSize:18}}>🐠</span>
              </div>
              <div style={{color:"#e0f7fa",fontSize:11,marginTop:3,fontWeight:600,letterSpacing:".05em"}}>PLATAFORMA DE REPRODUCAO E MORFOFISIOLOGIA · PANTANAL</div>
            </div>
          </div>
          <div style={{display:"flex",gap:8,alignItems:"center",flexWrap:"wrap"}}>
            <button onClick={gerarPlanilhaModelo} style={{...BS("#ffffff"),color:"#0077b6",padding:"8px 14px",fontSize:11.5,borderRadius:9,display:"flex",alignItems:"center",gap:5}}>
              📋 Baixar Planilha Modelo
            </button>
            <button onClick={()=>fileRef.current.click()} style={{...BS(loading?"#64748b":"#2d9e5f"),padding:"8px 14px",fontSize:11.5,borderRadius:9,display:"flex",alignItems:"center",gap:5}}>
              {loading?"⏳ Processando...":"📂 Carregar Minha Planilha"}
            </button>
            <input ref={fileRef} type="file" accept=".xlsx" style={{display:"none"}} onChange={e=>{handleFile(e.target.files[0]);e.target.value="";}}/>
          </div>
        </div>
      </div>

      {/* WELCOME STATE */}
      {!data&&(
        <div style={{maxWidth:700,margin:"48px auto",padding:"0 20px",position:"relative",zIndex:5}}>
          <div
            onDragOver={e=>{e.preventDefault();setDrag(true);}}
            onDragLeave={()=>setDrag(false)}
            onDrop={e=>{e.preventDefault();setDrag(false);handleFile(e.dataTransfer.files[0]);}}
            onClick={()=>fileRef.current.click()}
            style={{border:`2.5px dashed ${drag?"#0077b6":"#90caf9"}`,borderRadius:18,padding:"48px 32px",cursor:"pointer",background:drag?"#eff6ff":"#fff",display:"flex",flexDirection:"column",alignItems:"center",gap:14,textAlign:"center",transition:"all .2s",boxShadow:"0 4px 24px rgba(0,80,160,.08)"}}>
            <div style={{fontSize:56}}>{loading?"⏳":"📂"}</div>
            <div style={{fontWeight:900,color:"#1e3a5f",fontSize:17}}>{loading?"Processando sua planilha...":"Arraste sua planilha aqui"}</div>
            <div style={{fontSize:12.5,color:"#64748b",maxWidth:420,lineHeight:1.6}}>
              {loading
                ?"Aguarde enquanto os dados sao carregados e os graficos sao gerados."
                :"Faca o upload do arquivo .xlsx no formato padrao FishRep. Todos os graficos, tabelas e analises serao gerados automaticamente."}
            </div>
            {!loading&&<div style={{background:"#0077b6",color:"#fff",borderRadius:9,padding:"10px 20px",fontSize:12.5,fontWeight:700,marginTop:4}}>Selecionar arquivo .xlsx</div>}
          </div>

          <div style={{marginTop:20,background:"#fff",borderRadius:14,padding:"20px 24px",boxShadow:"0 2px 14px rgba(0,80,160,.07)"}}>
            <div style={{fontWeight:800,color:"#1e3a5f",fontSize:14,marginBottom:10}}>Como usar o FishRep</div>
            {[
              {n:"1",ic:"📋",t:"Baixe a planilha modelo",d:"Clique em 'Baixar Planilha Modelo' no cabecalho. O arquivo .xlsx ja vem com a estrutura correta de 7 abas e dados de exemplo do Pantanal."},
              {n:"2",ic:"✏️",t:"Preencha com seus dados",d:"Substitua os dados de exemplo pelos seus proprios dados de campo, mantendo exatamente a estrutura das colunas e abas."},
              {n:"3",ic:"📂",t:"Carregue e analise",d:"Arraste o arquivo aqui ou clique para selecionar. O FishRep gera todos os graficos, tabelas e analises estatisticas automaticamente."},
            ].map((s,i)=>(
              <div key={i} style={{display:"flex",gap:12,alignItems:"flex-start",marginBottom:14}}>
                <div style={{width:26,height:26,borderRadius:"50%",background:"#0077b6",color:"#fff",display:"flex",alignItems:"center",justifyContent:"center",fontSize:12,fontWeight:900,flexShrink:0}}>{s.n}</div>
                <div>
                  <div style={{fontWeight:700,color:"#1e3a5f",fontSize:12,marginBottom:2}}>{s.ic} {s.t}</div>
                  <div style={{fontSize:11.5,color:"#64748b",lineHeight:1.55}}>{s.d}</div>
                </div>
              </div>
            ))}
            <div style={{marginTop:6,paddingTop:12,borderTop:"1px solid #f1f5f9"}}>
              <div style={{fontWeight:700,color:"#1e3a5f",fontSize:12,marginBottom:8}}>Abas esperadas na planilha:</div>
              <div style={{display:"flex",gap:6,flexWrap:"wrap"}}>
                {["T1_RGS_Order_Month","T2_RGS_Species","T3_RGS_Sex_Month","T4_GonadStage_Freq","T5_Environmental_Month","T6_Statistical_Tests","T7_Sample_Overview"].map(t=>(
                  <span key={t} style={{background:"#eff6ff",color:"#1d4ed8",border:"1px solid #bfdbfe",borderRadius:6,padding:"3px 8px",fontSize:10,fontWeight:700}}>{t}</span>
                ))}
              </div>
            </div>
          </div>
        </div>
      )}

      {/* LOADED STATE */}
      {data&&(
        <>
          <div style={{background:"#fff",boxShadow:"0 2px 8px rgba(0,80,160,.07)",position:"sticky",top:0,zIndex:20}}>
            <div style={{maxWidth:1100,margin:"0 auto",padding:"0 16px",overflowX:"auto",display:"flex"}}>
              {TABS.map((t,i)=>(
                <button key={i} onClick={()=>setTab(i)} style={{background:"none",border:"none",cursor:"pointer",padding:"12px 14px",fontSize:12,fontWeight:700,color:tab===i?"#0077b6":"#64748b",borderBottom:tab===i?"3px solid #0077b6":"3px solid transparent",whiteSpace:"nowrap",display:"flex",alignItems:"center",gap:5,transition:"color .2s"}}>
                  {t.icon} {t.label}
                </button>
              ))}
              <button onClick={()=>{setData(null);setTab(0);}} style={{marginLeft:"auto",background:"none",border:"none",cursor:"pointer",padding:"12px 10px",fontSize:11,color:"#94a3b8",fontWeight:600,whiteSpace:"nowrap"}}>↩ Nova planilha</button>
            </div>
          </div>
          <div style={{maxWidth:1100,margin:"0 auto",padding:"20px 16px 40px",position:"relative",zIndex:5}} className="fu" key={tab}>
            {tab===0&&<TabOverview d={data}/>}
            {tab===1&&<TabRGSMensal d={data}/>}
            {tab===2&&<TabRGSEspecies d={data}/>}
            {tab===3&&<TabSexo d={data}/>}
            {tab===4&&<TabGonadas d={data}/>}
            {tab===5&&<TabAmbiental d={data}/>}
            {tab===6&&<TabEstatistica d={data}/>}
          </div>
        </>
      )}

      <div style={{textAlign:"center",padding:"12px",fontSize:10.5,color:"#94a3b8",position:"relative",zIndex:5}}>
        🐟 FishRep · Plataforma de Reproducao e Morfofisiologia de Peixes · Pantanal Norte 🐠
      </div>
    </div>
  );
}
