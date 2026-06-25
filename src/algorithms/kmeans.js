// ─── K-MEANS CLUSTERING ──────────────────────────────────────────────────────
function kDist(a,b,mL,mW,mH){
  const dl=(a.L-b.L)/Math.max(1,mL),dw=(a.W-b.W)/Math.max(1,mW),dh=(a.H-b.H)/Math.max(1,mH);
  return dl*dl+dw*dw+dh*dh; // squared distance (no sqrt needed for comparisons)
}

function runKMeans(points,k,maxIter=120){
  // points = [{L,W,H,qty,weight,name}, ...]
  if(!points.length||k<1) return[];
  k=Math.min(k,points.length);

  const mL=Math.max(...points.map(p=>p.L))||1;
  const mW=Math.max(...points.map(p=>p.W))||1;
  const mH=Math.max(...points.map(p=>p.H))||1;

  // K-means++ initialisation: spread starting centroids
  const centroids=[{...points[Math.floor(Math.random()*points.length)]}];
  while(centroids.length<k){
    // Pick next centroid with probability proportional to squared distance from nearest centroid
    const dists=points.map(p=>Math.min(...centroids.map(c=>kDist(p,c,mL,mW,mH))));
    const total=dists.reduce((a,b)=>a+b,0)||1;
    let r=Math.random()*total;
    for(let j=0;j<points.length;j++){r-=dists[j];if(r<=0){centroids.push({...points[j]});break;}}
    if(centroids.length<k) centroids.push({...points[points.length-1]}); // safety
  }

  const assign=new Int32Array(points.length);

  for(let iter=0;iter<maxIter;iter++){
    // Assignment step
    let changed=false;
    for(let i=0;i<points.length;i++){
      let best=0,bestD=Infinity;
      for(let j=0;j<k;j++){const d=kDist(points[i],centroids[j],mL,mW,mH);if(d<bestD){bestD=d;best=j;}}
      if(assign[i]!==best){assign[i]=best;changed=true;}
    }
    if(!changed) break;

    // Update step — weighted centroid (weight by qty)
    for(let j=0;j<k;j++){
      let sL=0,sW=0,sH=0,sQ=0;
      for(let i=0;i<points.length;i++){
        if(assign[i]!==j) continue;
        const q=points[i].qty||1;sL+=points[i].L*q;sW+=points[i].W*q;sH+=points[i].H*q;sQ+=q;
      }
      if(sQ>0){centroids[j]={L:sL/sQ,W:sW/sQ,H:sH/sQ};}
    }
  }

  // Build result groups
  const groups=[];
  for(let j=0;j<k;j++){
    const members=points.filter((_,i)=>assign[i]===j);
    if(!members.length) continue;
    const totalQty=members.reduce((s,p)=>s+(p.qty||1),0);
    const totalWt=members.reduce((s,p)=>s+(p.weight||0)*(p.qty||1),0);
    // Representative box: weighted centroid rounded to nearest mm
    const repL=Math.max(1,Math.round(members.reduce((s,p)=>s+p.L*(p.qty||1),0)/totalQty));
    const repW=Math.max(1,Math.round(members.reduce((s,p)=>s+p.W*(p.qty||1),0)/totalQty));
    const repH=Math.max(1,Math.round(members.reduce((s,p)=>s+p.H*(p.qty||1),0)/totalQty));
    const avgWt=totalWt/totalQty;
    // Accuracy: avg distance from centroid to members (as % of container diagonal)
    const avgErr=members.reduce((s,p)=>{
      const dl=Math.abs(p.L-repL)/mL,dw=Math.abs(p.W-repW)/mW,dh=Math.abs(p.H-repH)/mH;
      return s+Math.sqrt(dl*dl+dw*dw+dh*dh);
    },0)/members.length;
    groups.push({id:j+1,repL,repW,repH,totalQty,avgWt,skuCount:members.length,members,
      accuracy:Math.max(0,100-Math.round(avgErr*100))});
  }

  // Sort groups by representative box volume ascending
  groups.sort((a,b)=>a.repL*a.repW*a.repH-b.repL*b.repW*b.repH);
  groups.forEach((g,i)=>{g.id=i+1;g.name=`Group ${i+1}`;});
  return groups;
}

// priority: "strict" = honour ratio exactly (may waste space)
//           "balanced" = stay close to ratio but fill gaps (default)

export { kDist, runKMeans };
