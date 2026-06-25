// ─── MULTI-SKU ALGORITHMS ────────────────────────────────────────────────────
import { bestFitDetailed } from './packing.js';

export const MULTI_PALETTE=[0xbe185d,0x0ea5e9,0x16a34a,0xd97706,0x7c3aed,0x0891b2,0xf97316,0x374151];
export const MULTI_LABELS=['#be185d','#0ea5e9','#16a34a','#d97706','#7c3aed','#0891b2','#f97316','#374151'];

function calcMultiSKU(cL,cW,cH,skus,opt={}){
  // skus = [{name,L,W,H,targetQty}, ...]
  // Divides container into N strips (one per SKU) along each axis.
  // Tries L, W, H cuts and returns the axis with best total utilization.
  const valid=skus.filter(s=>s.L>0&&s.W>0&&s.H>0&&s.targetQty>0);
  if(!valid.length) return null;

  const cVol=cL*cW*cH;
  // Volume each SKU needs to hit its target qty
  const vNeed=valid.map(s=>s.targetQty*s.L*s.W*s.H);
  const vTotal=vNeed.reduce((a,b)=>a+b,0)||1;
  const fracs=vNeed.map(v=>v/vTotal);

  // Try dividing along one axis, proportional to volume fractions
  function tryAxis(axLen,mkReg){
    const regs=[];let off=0;
    valid.forEach((sku,i)=>{
      const isLast=i===valid.length-1;
      // Snap allocation to nearest box dimension on that axis for cleaner packing
      const snap=bestFitDetailed(
        mkReg(0,axLen).L,mkReg(0,axLen).W,mkReg(0,axLen).H,
        sku.L,sku.W,sku.H,opt);
      const boxOnAxis=axLen===cL?snap.boxL:axLen===cW?snap.boxW:snap.boxH;
      let len=isLast?axLen-off:Math.round(fracs[i]*axLen);
      if(boxOnAxis>0&&!isLast){
        const snapped=Math.max(boxOnAxis,Math.round(len/boxOnAxis)*boxOnAxis);
        len=Math.min(snapped,axLen-off-(valid.length-i-1)*Math.max(1,boxOnAxis));
      }
      len=Math.max(1,Math.min(len,axLen-off));
      const r=mkReg(off,len);
      const det=bestFitDetailed(r.L,r.W,r.H,sku.L,sku.W,sku.H,opt);
      regs.push({...sku,fitted:det.count,target:sku.targetQty,
        fillRate:Math.min(1,det.count/Math.max(1,sku.targetQty)),
        det,off:r.off,regionDims:{L:r.L,W:r.W,H:r.H},
        skuVol:(det.count*sku.L*sku.W*sku.H)/(r.L*r.W*r.H)});
      off+=len;
    });
    const fVol=regs.reduce((s,r)=>s+r.fitted*r.L*r.W*r.H,0);
    const avgFill=regs.reduce((s,r)=>s+r.fillRate,0)/regs.length;
    const score=fVol/cVol+avgFill*0.3;
    return{regs,score,total:regs.reduce((s,r)=>s+r.fitted,0),volUtil:fVol/cVol};
  }

  const attempts=[
    tryAxis(cL,(off,len)=>({L:len,W:cW,H:cH,off:{x:off,y:0,z:0}})),
    tryAxis(cW,(off,len)=>({L:cL,W:len,H:cH,off:{x:0,y:off,z:0}})),
    tryAxis(cH,(off,len)=>({L:cL,W:cW,H:len,off:{x:0,y:0,z:off}})),
  ];
  const best=attempts.reduce((a,b)=>a.score>b.score?a:b);

  // Apply weight constraint: if total fitted weight exceeds maxWeight,
  // scale back each region proportionally by weight budget fraction
  if(opt.maxWeight>0){
    const totalWt=best.regs.reduce((s,r)=>s+(r.weight||0)*r.fitted,0);
    if(totalWt>opt.maxWeight&&totalWt>0){
      const scale=opt.maxWeight/totalWt;
      best.regs.forEach(r=>{
        r.fitted=Math.floor(r.fitted*scale);
        r.fillRate=Math.min(1,r.fitted/Math.max(1,r.targetQty));
        r.weightConstrained=true;
      });
      best.total=best.regs.reduce((s,r)=>s+r.fitted,0);
    }
  }

  return{...best,cL,cW,cH,regions:best.regs};
}

// ─── MULTI-SKU SHIPMENT PLANNER ───────────────────────────────────────────────
// Runs calcMultiSKU repeatedly until all SKU quantities are shipped.
// Each iteration = one container. Returns array of containers + summary.
function calcMultiSKUShipment(cL,cW,cH,maxWt,skus,opt={}){
  const optW={...opt,maxWeight:maxWt||0};
  // Track shipped quantities by SKU name
  const shipped=Object.fromEntries(skus.map(s=>[s.name,0]));
  const containers=[];const MAX_CONTAINERS=200;

  while(containers.length<MAX_CONTAINERS){
    // Build target array from remaining quantities
    const targets=skus
      .map(s=>({...s,targetQty:s.qty-(shipped[s.name]||0)}))
      .filter(s=>s.targetQty>0);
    if(!targets.length) break;

    const res=calcMultiSKU(cL,cW,cH,targets,optW);
    if(!res||res.total===0) break;

    // Progress check — prevent infinite loop
    const before=Object.values(shipped).reduce((a,b)=>a+b,0);
    res.regions.forEach(r=>{
      if(r.fitted>0) shipped[r.name]=(shipped[r.name]||0)+r.fitted;
    });
    const after=Object.values(shipped).reduce((a,b)=>a+b,0);
    if(after===before) break;

    containers.push({
      id:containers.length+1,
      regions:res.regions,
      total:res.total,
      volUtil:res.volUtil,
      totalWeight:res.regions.reduce((s,r)=>s+r.fitted*(r.weight||0),0),
      cL,cW,cH,
    });
  }

  const totalShipped=Object.values(shipped).reduce((a,b)=>a+b,0);
  const totalOrdered=skus.reduce((s,sk)=>s+sk.qty,0);
  const avgUtil=containers.length
    ?containers.reduce((s,c)=>s+c.volUtil,0)/containers.length:0;
  return{containers,totalContainers:containers.length,
    totalShipped,totalOrdered,avgUtil,shipped,skus};
}

export { calcMultiSKU, calcMultiSKUShipment };
