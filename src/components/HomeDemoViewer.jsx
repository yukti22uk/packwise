// ─── INTERACTIVE 3D HOMEPAGE DEMO ────────────────────────────────────────────
import { useState, useRef, useEffect } from 'react';
import * as THREE from 'three';
import { calcMixedDetailed } from '../algorithms/packing.js';
import { fmtN } from '../algorithms/utils.js';
// ─── INTERACTIVE 3D HOMEPAGE DEMO ────────────────────────────────────────────
const HOME_PRESETS=[
  {name:"Tata Ace",  box:[300,200,150],container:[2100,1525,1600]},
  {name:"20ft ISO",  box:[500,400,300],container:[5900,2350,2390]},
  {name:"32ft SXL",  box:[600,400,300],container:[9750,2350,2700]},
  {name:"Pallet",    box:[300,250,200],container:[1200,1000,1400]},
];

function HomeDemoViewer(){
  const mountRef=useRef(null);const cleanRef=useRef(null);
  const[preset,setPreset]=useState(1);const[result,setResult]=useState(null);
  useEffect(()=>{
    const p=HOME_PRESETS[preset];
    const r=calcMixedDetailed(p.container[0],p.container[1],p.container[2],p.box[0],p.box[1],p.box[2]);
    setResult({...r,volUtil:(r.total*p.box[0]*p.box[1]*p.box[2])/(p.container[0]*p.container[1]*p.container[2]),preset:p});
  },[preset]);
  useEffect(()=>{
    if(!result||!mountRef.current) return;
    if(cleanRef.current){cleanRef.current();cleanRef.current=null;}
    const t=setTimeout(()=>{
      const el=mountRef.current;if(!el) return;
      const W=el.clientWidth||700,H=340;
      const{cL,cW,cH,nx,ny,nz,boxL,boxW,boxH,leftover1:l1,leftover2:l2,leftover3:l3}=result;
      const scene=new THREE.Scene();scene.background=new THREE.Color(0x0f172a);
      const camera=new THREE.PerspectiveCamera(45,W/H,0.01,100000);
      const renderer=new THREE.WebGLRenderer({antialias:true});
      renderer.setSize(W,H);renderer.setPixelRatio(Math.min(window.devicePixelRatio,2));
      el.appendChild(renderer.domElement);renderer.domElement.style.display="block";
      scene.add(new THREE.AmbientLight(0xffffff,0.75));
      const d1=new THREE.DirectionalLight(0xffffff,0.7);d1.position.set(5,8,5);scene.add(d1);
      const d2=new THREE.DirectionalLight(0xff88aa,0.25);d2.position.set(-3,-2,-3);scene.add(d2);
      const regs=[
        {col:0xbe185d,ox:0,oy:0,oz:0,rnx:nx,rny:ny,rnz:nz,bL:boxL,bW:boxW,bH:boxH},
        {col:0x374151,ox:l1.offX,oy:l1.offY,oz:l1.offZ,rnx:l1.nx,rny:l1.ny,rnz:l1.nz,bL:l1.boxL,bW:l1.boxW,bH:l1.boxH},
        {col:0x4b5563,ox:l2.offX,oy:l2.offY,oz:l2.offZ,rnx:l2.nx,rny:l2.ny,rnz:l2.nz,bL:l2.boxL,bW:l2.boxW,bH:l2.boxH},
        {col:0x6b7280,ox:l3.offX,oy:l3.offY,oz:l3.offZ,rnx:l3.nx,rny:l3.ny,rnz:l3.nz,bL:l3.boxL,bW:l3.boxW,bH:l3.boxH},
      ];
      const tot=regs.reduce((s,r)=>s+(r.rnx||0)*(r.rny||0)*(r.rnz||0),0)||1;
      const MAX_D=900;
      regs.forEach(r=>{
        if(!r.rnx||!r.rny||!r.rnz||!r.bL||!r.bW||!r.bH) return;
        const total=r.rnx*r.rny*r.rnz,cap=Math.max(1,Math.round(MAX_D*(total/tot)));
        const stride=total>cap?total/cap:1;const pos=[];let idx=0,next=0;
        for(let iz=0;iz<r.rnz;iz++)for(let iy=0;iy<r.rny;iy++)for(let ix=0;ix<r.rnx;ix++){
          const e=(ix===0||ix===r.rnx-1||iy===0||iy===r.rny-1||iz===0||iz===r.rnz-1);
          if(e||idx>=next){pos.push([(r.ox||0)+ix*r.bL+r.bL/2,(r.oz||0)+iz*r.bH+r.bH/2,(r.oy||0)+iy*r.bW+r.bW/2]);
            if(idx>=next)next+=stride;}idx++;}
        const mesh=new THREE.InstancedMesh(new THREE.BoxGeometry(r.bL*0.86,r.bH*0.86,r.bW*0.86),
          new THREE.MeshPhongMaterial({color:r.col,shininess:40,transparent:true,opacity:0.96}),pos.length);
        mesh.count=pos.length;const dummy=new THREE.Object3D();
        pos.forEach(([x,y,z],i)=>{dummy.position.set(x,y,z);dummy.updateMatrix();mesh.setMatrixAt(i,dummy.matrix);});
        mesh.instanceMatrix.needsUpdate=true;scene.add(mesh);});
      const cw=new THREE.LineSegments(new THREE.EdgesGeometry(new THREE.BoxGeometry(cL,cH,cW)),
        new THREE.LineBasicMaterial({color:0xbe185d}));
      cw.position.set(cL/2,cH/2,cW/2);scene.add(cw);
      const grid=new THREE.GridHelper(Math.max(cL,cW)*2.5,10,0x1e293b,0x1a2030);
      grid.position.set(cL/2,-1,cW/2);scene.add(grid);
      const center=new THREE.Vector3(cL/2,cH/2,cW/2);
      const diag=Math.sqrt(cL*cL+cW*cW+cH*cH)||10;
      let radius=diag*2.1,theta=Math.PI*0.3,phi=Math.PI*0.28,drag=false,prev={x:0,y:0};
      const cv=renderer.domElement;cv.style.cursor="grab";
      const onD=(e)=>{drag=true;prev={x:e.clientX,y:e.clientY};cv.style.cursor="grabbing";};
      const onM=(e)=>{if(!drag)return;theta-=(e.clientX-prev.x)*0.006;
        phi=Math.max(0.05,Math.min(1.5,phi+(e.clientY-prev.y)*0.006));prev={x:e.clientX,y:e.clientY};};
      const onU=()=>{drag=false;cv.style.cursor="grab";};
      const onW=(e)=>{e.preventDefault();radius=Math.max(diag*0.5,Math.min(diag*5,radius+e.deltaY*0.4));};
      cv.addEventListener("mousedown",onD);window.addEventListener("mousemove",onM);
      window.addEventListener("mouseup",onU);cv.addEventListener("wheel",onW,{passive:false});
      const onRz=()=>{const nW=el.clientWidth||700;renderer.setSize(nW,H);camera.aspect=nW/H;camera.updateProjectionMatrix();};
      window.addEventListener("resize",onRz);
      let animId;
      const loop=()=>{animId=requestAnimationFrame(loop);
        if(!drag)theta+=0.003;
        camera.position.set(center.x+radius*Math.sin(phi)*Math.sin(theta),
          center.y+radius*Math.cos(phi),center.z+radius*Math.sin(phi)*Math.cos(theta));
        camera.lookAt(center);renderer.render(scene,camera);};loop();
      cleanRef.current=()=>{cancelAnimationFrame(animId);
        cv.removeEventListener("mousedown",onD);window.removeEventListener("mousemove",onM);
        window.removeEventListener("mouseup",onU);cv.removeEventListener("wheel",onW);
        window.removeEventListener("resize",onRz);
        if(el.contains(cv))el.removeChild(cv);renderer.dispose();};
    },120);
    return()=>{clearTimeout(t);if(cleanRef.current){cleanRef.current();cleanRef.current=null;}};
  },[result]);
  return(
    <div style={{background:"#0f172a",padding:"72px 24px",borderTop:"1px solid #1e293b"}}>
      <div style={{maxWidth:"1200px",margin:"0 auto"}}>
        <FadeIn style={{textAlign:"center",marginBottom:"32px"}}>
          <div style={{fontSize:"12px",fontWeight:"700",color:"#f9a8d4",letterSpacing:"0.1em",
            textTransform:"uppercase",marginBottom:"10px"}}>Live interactive demo</div>
          <h2 style={{fontSize:"36px",fontWeight:"900",color:"#fff",margin:"0 0 8px",letterSpacing:"-0.02em"}}>
            See the 3D model — before you sign up
          </h2>
          <p style={{fontSize:"15px",color:"#94a3b8",margin:0}}>
            Real packing algorithm. Real 3D model. Drag to rotate, scroll to zoom.
          </p>
        </FadeIn>
        <div style={{display:"flex",gap:"8px",justifyContent:"center",marginBottom:"20px",flexWrap:"wrap"}}>
          {HOME_PRESETS.map((p,i)=>(
            <button key={i} onClick={()=>setPreset(i)} style={{padding:"8px 20px",
              border:`1.5px solid ${preset===i?"#be185d":"#1e293b"}`,
              background:preset===i?"#be185d":"transparent",
              color:preset===i?"#fff":"#9ca3af",borderRadius:"99px",cursor:"pointer",
              fontWeight:"700",fontSize:"13px",fontFamily:"inherit",transition:"all 0.15s ease"}}>
              {p.name}</button>
          ))}
        </div>
        <div ref={mountRef} style={{width:"100%",height:"340px",borderRadius:"14px",
          overflow:"hidden",border:"1px solid #1e293b",background:"#0f172a",
          boxShadow:"0 0 60px rgba(190,24,93,0.08)"}}/>
        {result&&(
          <FadeIn style={{display:"flex",gap:"16px",justifyContent:"center",marginTop:"20px",flexWrap:"wrap"}}>
            {[["📦","Boxes Fit",result.total.toLocaleString()],
              ["📐","Space Used",(result.volUtil*100).toFixed(1)+"%"],
              ["🔄","Orientation",result.orient],
            ].map(([icon,label,val])=>(
              <div key={label} style={{textAlign:"center",padding:"14px 24px",background:"#1e293b",
                borderRadius:"10px",border:"1px solid #263040",minWidth:"140px"}}>
                <div style={{fontSize:"22px",fontWeight:"900",color:"#fff",letterSpacing:"-0.02em"}}>{val}</div>
                <div style={{fontSize:"11px",color:"#6b7280",marginTop:"4px",fontWeight:"600",
                  textTransform:"uppercase",letterSpacing:"0.06em"}}>{icon} {label}</div>
              </div>
            ))}
          </FadeIn>
        )}
        <div style={{textAlign:"center",marginTop:"20px",fontSize:"12px",color:"#4b5563"}}>
          Drag to rotate · Scroll to zoom · Switch presets to see different containers
        </div>
      </div>
    </div>
  );
}

// ── Nav ──

export default HomeDemoViewer;
