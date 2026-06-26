// ─── CORE PACKING ALGORITHMS ─────────────────────────────────────────────────
import { fmtN } from './utils.js';
function getPerms(sl,sw,sh,lockHeight){
  if(lockHeight) return [[sl,sw,sh],[sw,sl,sh]];
  return [[sl,sw,sh],[sl,sh,sw],[sw,sl,sh],[sw,sh,sl],[sh,sl,sw],[sh,sw,sl]];
}
function bestFitDetailed(L,W,H,sl,sw,sh,opt={}){
  if(L<=0||W<=0||H<=0) return{count:0,nx:0,ny:0,nz:0,boxL:sl,boxW:sw,boxH:sh};
  const perms=getPerms(sl,sw,sh,opt.lockHeight);
  let best=0,bx=[sl,sw,sh],bn=[0,0,0];
  for(const[a,b,c]of perms){
    let nx=Math.floor(L/a),ny=Math.floor(W/b),nz=Math.floor(H/c);
    if(opt.noStack) nz=Math.min(nz,1);
    if(opt.maxStack&&opt.maxStack>0) nz=Math.min(nz,opt.maxStack);
    const cnt=nx*ny*nz;
    if(cnt>best){best=cnt;bx=[a,b,c];bn=[nx,ny,nz];}
  }
  return{count:best,nx:bn[0],ny:bn[1],nz:bn[2],boxL:bx[0],boxW:bx[1],boxH:bx[2]};
}
function calcMixedDetailed(cL,cW,cH,sl,sw,sh,opt={}){
  const perms=getPerms(sl,sw,sh,opt.lockHeight);
  let best=-1,R={};
  for(const[a,b,c]of perms){
    let nx=Math.floor(cL/a),ny=Math.floor(cW/b),nz=Math.floor(cH/c);
    if(opt.noStack) nz=Math.min(nz,1);
    if(opt.maxStack&&opt.maxStack>0) nz=Math.min(nz,opt.maxStack);
    const uL=nx*a,uW=ny*b,uH=nz*c;
    const l1=bestFitDetailed(cL-uL,cW,cH,sl,sw,sh,opt);
    const l2=bestFitDetailed(uL,cW-uW,cH,sl,sw,sh,opt);
    const l3=(opt.noStack||(opt.maxStack&&opt.maxStack>0))?{count:0,nx:0,ny:0,nz:0,boxL:sl,boxW:sw,boxH:sh}
              :bestFitDetailed(uL,uW,cH-uH,sl,sw,sh,opt);
    const tot=nx*ny*nz+l1.count+l2.count+l3.count;
    if(tot>best){best=tot;
      R={total:tot,nx,ny,nz,boxL:a,boxW:b,boxH:c,usedL:uL,usedW:uW,usedH:uH,cL,cW,cH,sl,sw,sh,
        leftover1:{...l1,offX:uL,offY:0,offZ:0},
        leftover2:{...l2,offX:0,offY:uW,offZ:0},
        leftover3:{...l3,offX:0,offY:0,offZ:uH},
        orient:`${fmtN(a)}×${fmtN(b)}×${fmtN(c)}`};}
  }
  return R;
}
function effectivePerContainer(cL,cW,cH,sku,opt={}){
  const r=calcMixedDetailed(cL,cW,cH,sku.L,sku.W,sku.H,opt);
  let volQty=r.total,wtQty=Infinity;
  if(sku.weight>0 && opt.maxWeight>0) wtQty=Math.floor(opt.maxWeight/sku.weight);
  const eff=Math.min(volQty,wtQty);
  return{...r,volQty,wtQty:wtQty===Infinity?null:wtQty,eff,constraint:wtQty<volQty?"Weight":"Volume"};
}
function calcMixed(cL,cW,cH,sl,sw,sh){const r=calcMixedDetailed(cL,cW,cH,sl,sw,sh);return{total:r.total,orient:r.orient};}


export { getPerms, bestFitDetailed, calcMixedDetailed, effectivePerContainer, calcMixed };
