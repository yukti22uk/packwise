// ─── 2D ENGINEERING VIEWS ────────────────────────────────────────────────────
import { useState, useRef, useEffect } from 'react';
import { fmtN } from '../algorithms/utils.js';
const RC={main:{fill:"#93c5fd",stroke:"#1d4ed8",label:"Main Grid"},l1:{fill:"#fdba74",stroke:"#ea580c",label:"Leftover 1 (side)"},
  l2:{fill:"#86efac",stroke:"#16a34a",label:"Leftover 2 (front)"},l3:{fill:"#d8b4fe",stroke:"#7c3aed",label:"Leftover 3 (top)"}};
function BoxGrid({offX,offY,nx,ny,bW,bH,sc,color,dimLabel}){
  if(!nx||!ny||!bW||!bH) return null;const cells=[];const total=nx*ny,skip=total>2000?Math.ceil(total/2000):1;
  for(let iy=0;iy<ny;iy++)for(let ix=0;ix<nx;ix++){if((iy*nx+ix)%skip!==0&&!(ix===0&&iy===0)) continue;
    const x=(offX+ix*bW)*sc+1,y=(offY+iy*bH)*sc+1,w=bW*sc-1,h=bH*sc-1;
    cells.push(<g key={`${ix}-${iy}`}><rect x={x} y={y} width={Math.max(0,w)} height={Math.max(0,h)} fill={color.fill} stroke={color.stroke} strokeWidth="0.6"/>
      {ix===0&&iy===0&&w>18&&h>10&&<text x={x+w/2} y={y+h/2} textAnchor="middle" dominantBaseline="middle"
        fontSize={Math.max(7,Math.min(10,(Math.min(w,h)-4)/2.2))} fill="#111">{dimLabel}</text>}</g>);}
  return<>{cells}</>;}
function SvgDefs(){return(<defs><marker id="arr" markerWidth="4" markerHeight="4" refX="2" refY="2" orient="auto"><path d="M0,0 L4,2 L0,4 Z" fill="#475569"/></marker></defs>);}
function DimArrow({x1,y1,x2,y2,label,pos="top"}){const mx=(x1+x2)/2,my=(y1+y2)/2;
  return(<g><line x1={x1} y1={y1} x2={x2} y2={y2} stroke="#475569" strokeWidth="1" markerEnd="url(#arr)" markerStart="url(#arr)"/>
    <text x={mx+(pos==="left"?-6:0)} y={my+(pos==="top"?-10:0)} textAnchor="middle" fontSize="9" fill="#475569" fontWeight="600">{label}</text></g>);}
const VL={viewTitle:{fontSize:"13px",fontWeight:"700",color:"#1a2332",marginBottom:"8px",textAlign:"center"},note:{fontSize:"10px",color:"#9ca3af",marginTop:"4px",fontStyle:"italic",textAlign:"center"}};
function TopView2D({result}){const{cL,cW,nx,ny,boxL,boxW,usedL,usedW,leftover1:l1,leftover2:l2}=result;
  const sc=Math.min(300/cL,300/cW,20),W=cL*sc,H=cW*sc,P=28;
  return(<div style={{textAlign:"center"}}><div style={VL.viewTitle}>Top View (L × W)</div>
    <svg width={W+P+10} height={H+P+10} style={{display:"block",margin:"0 auto"}}><SvgDefs/>
      <rect x={P} y={P} width={W} height={H} fill="#f8fafc" stroke="#1e293b" strokeWidth="2"/>
      <g transform={`translate(${P},${P})`}>
        <BoxGrid offX={0} offY={0} nx={nx} ny={ny} bW={boxL} bH={boxW} sc={sc} color={RC.main} dimLabel={`${fmtN(boxL)}×${fmtN(boxW)}`}/>
        {l1.count>0&&<BoxGrid offX={usedL} offY={0} nx={l1.nx} ny={l1.ny} bW={l1.boxL} bH={l1.boxW} sc={sc} color={RC.l1} dimLabel={`${fmtN(l1.boxL)}×${fmtN(l1.boxW)}`}/>}
        {l2.count>0&&<BoxGrid offX={0} offY={usedW} nx={l2.nx} ny={l2.ny} bW={l2.boxL} bH={l2.boxW} sc={sc} color={RC.l2} dimLabel={`${fmtN(l2.boxL)}×${fmtN(l2.boxW)}`}/>}
        {usedL<cL&&<line x1={usedL*sc} y1={0} x2={usedL*sc} y2={H} stroke="#64748b" strokeWidth="1" strokeDasharray="3,2"/>}
        {usedW<cW&&<line x1={0} y1={usedW*sc} x2={W} y2={usedW*sc} stroke="#64748b" strokeWidth="1" strokeDasharray="3,2"/>}
      </g>
      <DimArrow x1={P} y1={P-14} x2={P+W} y2={P-14} label={fmtN(cL)} pos="top"/>
      <DimArrow x1={P-16} y1={P} x2={P-16} y2={P+H} label={fmtN(cW)} pos="left"/>
    </svg><div style={VL.note}>L3 (top gap) not shown — see Side View</div></div>);}
function SideView2D({result}){const{cL,cH,nx,nz,boxL,boxH,usedL,usedH,leftover1:l1,leftover3:l3}=result;
  const sc=Math.min(300/cL,300/cH,20),W=cL*sc,H=cH*sc,P=28;
  return(<div style={{textAlign:"center"}}><div style={VL.viewTitle}>Side View (L × H)</div>
    <svg width={W+P+10} height={H+P+10} style={{display:"block",margin:"0 auto"}}><SvgDefs/>
      <rect x={P} y={P} width={W} height={H} fill="#f8fafc" stroke="#1e293b" strokeWidth="2"/>
      <g transform={`translate(${P},${P})`}>
        {Array.from({length:nz},(_,iz)=>Array.from({length:nx},(_,ix)=>{const x=ix*boxL*sc+1,y=H-(iz+1)*boxH*sc+1,w=boxL*sc-1,h=boxH*sc-1;
          return(<g key={`m-${iz}-${ix}`}><rect x={x} y={y} width={w} height={h} fill={RC.main.fill} stroke={RC.main.stroke} strokeWidth="0.6"/>
            {ix===0&&iz===0&&w>18&&h>10&&<text x={x+w/2} y={y+h/2} textAnchor="middle" dominantBaseline="middle" fontSize={Math.max(7,Math.min(10,(Math.min(w,h)-4)/2.2))} fill="#111">{fmtN(boxL)}×{fmtN(boxH)}</text>}</g>);})).flat()}
        {l1.count>0&&Array.from({length:l1.nz},(_,iz)=>Array.from({length:l1.nx},(_,ix)=>{const x=(usedL+ix*l1.boxL)*sc+1,y=H-(iz+1)*l1.boxH*sc+1,w=l1.boxL*sc-1,h=l1.boxH*sc-1;
          return(<g key={`l1-${iz}-${ix}`}><rect x={x} y={y} width={Math.max(0,w)} height={Math.max(0,h)} fill={RC.l1.fill} stroke={RC.l1.stroke} strokeWidth="0.6"/>
            {ix===0&&iz===0&&w>18&&h>10&&<text x={x+w/2} y={y+h/2} textAnchor="middle" dominantBaseline="middle" fontSize={Math.max(7,Math.min(10,(Math.min(w,h)-4)/2.2))} fill="#111">{fmtN(l1.boxL)}×{fmtN(l1.boxH)}</text>}</g>);})).flat()}
        {l3.count>0&&Array.from({length:l3.nz},(_,iz)=>Array.from({length:l3.nx},(_,ix)=>{const x=ix*l3.boxL*sc+1,y=H-(usedH+(iz+1)*l3.boxH)*sc+1,w=l3.boxL*sc-1,h=l3.boxH*sc-1;
          return(<g key={`l3-${iz}-${ix}`}><rect x={x} y={y} width={Math.max(0,w)} height={Math.max(0,h)} fill={RC.l3.fill} stroke={RC.l3.stroke} strokeWidth="0.6"/>
            {ix===0&&iz===0&&w>18&&h>10&&<text x={x+w/2} y={y+h/2} textAnchor="middle" dominantBaseline="middle" fontSize={Math.max(7,Math.min(10,(Math.min(w,h)-4)/2.2))} fill="#111">{fmtN(l3.boxL)}×{fmtN(l3.boxH)}</text>}</g>);})).flat()}
        {usedL<cL&&<line x1={usedL*sc} y1={0} x2={usedL*sc} y2={H} stroke="#64748b" strokeWidth="1" strokeDasharray="3,2"/>}
        {usedH<cH&&<line x1={0} y1={H-usedH*sc} x2={W} y2={H-usedH*sc} stroke="#64748b" strokeWidth="1" strokeDasharray="3,2"/>}
      </g>
      <DimArrow x1={P} y1={P-14} x2={P+W} y2={P-14} label={fmtN(cL)} pos="top"/>
      <DimArrow x1={P-16} y1={P} x2={P-16} y2={P+H} label={fmtN(cH)} pos="left"/>
    </svg><div style={VL.note}>L2 (front gap) not shown — see Top View</div></div>);}
function IsoView2D({result}){const{cL,cW,cH,nx,ny,nz,boxL,boxW,boxH,usedL,usedW,usedH,leftover1:l1,leftover2:l2,leftover3:l3}=result;
  const MAX=260;let sc=MAX/Math.max(cL,cW,cH);if(sc>14)sc=14;if(sc<0.1)sc=0.1;
  const c30=0.8660254;function ix(x,y){return(x-y)*sc*c30;}function iy(x,y,z){return(x+y)*sc*0.5-z*sc;}
  const corners=[[0,0,0],[cL,0,0],[0,cW,0],[cL,cW,0],[0,0,cH],[cL,0,cH],[0,cW,cH],[cL,cW,cH]];
  const xs=corners.map(([x,y])=>ix(x,y)),ys=corners.map(([x,y,z])=>iy(x,y,z));
  const minX=Math.min(...xs),minY=Math.min(...ys),svgW=Math.max(...xs)-minX+20,svgH=Math.max(...ys)-minY+20,ox=-minX+10,oy=-minY+10;
  function px(x,y){return ix(x,y)+ox;}function py(x,y,z){return iy(x,y,z)+oy;}
  function blk(x0,y0,z0,x1,y1,z1,tC,fC,rC){if(x0>=x1||y0>=y1||z0>=z1)return null;
    const T=`${px(x0,y0,z1).toFixed(1)},${py(x0,y0,z1).toFixed(1)} ${px(x1,y0,z1).toFixed(1)},${py(x1,y0,z1).toFixed(1)} ${px(x1,y1,z1).toFixed(1)},${py(x1,y1,z1).toFixed(1)} ${px(x0,y1,z1).toFixed(1)},${py(x0,y1,z1).toFixed(1)}`;
    const F=`${px(x0,y0,z0).toFixed(1)},${py(x0,y0,z0).toFixed(1)} ${px(x1,y0,z0).toFixed(1)},${py(x1,y0,z0).toFixed(1)} ${px(x1,y0,z1).toFixed(1)},${py(x1,y0,z1).toFixed(1)} ${px(x0,y0,z1).toFixed(1)},${py(x0,y0,z1).toFixed(1)}`;
    const R=`${px(x1,y0,z0).toFixed(1)},${py(x1,y0,z0).toFixed(1)} ${px(x1,y1,z0).toFixed(1)},${py(x1,y1,z0).toFixed(1)} ${px(x1,y1,z1).toFixed(1)},${py(x1,y1,z1).toFixed(1)} ${px(x1,y0,z1).toFixed(1)},${py(x1,y0,z1).toFixed(1)}`;
    return<g opacity={0.88}><polygon points={T} fill={tC} stroke="#1e293b" strokeWidth="0.7"/><polygon points={F} fill={fC} stroke="#1e293b" strokeWidth="0.7"/><polygon points={R} fill={rC} stroke="#1e293b" strokeWidth="0.7"/></g>;}
  function edge(x0,y0,z0,x1,y1,z1,dash=false){return<line x1={px(x0,y0).toFixed(1)} y1={py(x0,y0,z0).toFixed(1)} x2={px(x1,y1).toFixed(1)} y2={py(x1,y1,z1).toFixed(1)} stroke="#1e293b" strokeWidth="1.5" strokeDasharray={dash?"4,3":"none"}/>;}
  const mX=usedL,mY=usedW,mZ=usedH;
  const l1X=usedL+l1.nx*l1.boxL,l1Y=l1.ny*l1.boxW,l1Z=l1.nz*l1.boxH;
  const l2X=l2.nx*l2.boxL,l2Y=usedW+l2.ny*l2.boxW,l2Z=l2.nz*l2.boxH;
  const l3X=l3.nx*l3.boxL,l3Y=l3.ny*l3.boxW,l3Z=usedH+l3.nz*l3.boxH;
  return(<div style={{textAlign:"center"}}><div style={VL.viewTitle}>Isometric View</div>
    <svg width={svgW} height={svgH} style={{display:"block",margin:"0 auto"}}>
      {l2.count>0&&blk(0,mY,0,l2X,l2Y,l2Z,"#bbf7d0","#86efac","#4ade80")}
      {l3.count>0&&blk(0,0,mZ,l3X,l3Y,l3Z,"#e9d5ff","#c084fc","#a855f7")}
      {l1.count>0&&blk(mX,0,0,l1X,l1Y,l1Z,"#fed7aa","#fb923c","#f97316")}
      {nx>0&&ny>0&&nz>0&&blk(0,0,0,mX,mY,mZ,"#bfdbfe","#60a5fa","#3b82f6")}
      {edge(0,0,0,cL,0,0)}{edge(0,0,0,0,cW,0)}{edge(0,0,0,0,0,cH)}{edge(cL,0,0,cL,cW,0)}{edge(cL,0,0,cL,0,cH)}
      {edge(0,cW,0,cL,cW,0)}{edge(0,cW,0,0,cW,cH)}{edge(0,0,cH,cL,0,cH)}{edge(0,0,cH,0,cW,cH)}
      {edge(cL,cW,0,cL,cW,cH)}{edge(cL,0,cH,cL,cW,cH)}{edge(0,cW,cH,cL,cW,cH)}{edge(0,cW,0,0,cW,cH,true)}
    </svg>
    <div style={{display:"flex",gap:"10px",justifyContent:"center",flexWrap:"wrap",marginTop:"6px"}}>
      {[["#60a5fa","Main"],["#fb923c","L1"],["#86efac","L2"],["#c084fc","L3"]].map(([c,l])=>(
        <div key={l} style={{display:"flex",alignItems:"center",gap:"4px",fontSize:"11px",color:"#555"}}><div style={{width:"11px",height:"11px",background:c,borderRadius:"2px"}}/>{l}</div>))}</div></div>);}


export { TopView2D, SideView2D, IsoView2D };
