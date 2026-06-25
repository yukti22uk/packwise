// ─── THREE.JS 3D VIEWER ──────────────────────────────────────────────────────
import { useState, useRef, useEffect } from 'react';
import * as THREE from 'three';
// ─── 3D VIEWER ────────────────────────────────────────────────────────────────
const MAX_3D=1200;
function ThreeViewer({result,captureRef,regions3D}){
  const mountRef=useRef(null);const cleanRef=useRef(null);
  const regs=regions3D||[
    {col:0x3b82f6,ox:0,oy:0,oz:0,rnx:result.nx,rny:result.ny,rnz:result.nz,bL:result.boxL,bW:result.boxW,bH:result.boxH},
    {col:0xf97316,ox:result.leftover1.offX,oy:result.leftover1.offY,oz:result.leftover1.offZ,rnx:result.leftover1.nx,rny:result.leftover1.ny,rnz:result.leftover1.nz,bL:result.leftover1.boxL,bW:result.leftover1.boxW,bH:result.leftover1.boxH},
    {col:0x22c55e,ox:result.leftover2.offX,oy:result.leftover2.offY,oz:result.leftover2.offZ,rnx:result.leftover2.nx,rny:result.leftover2.ny,rnz:result.leftover2.nz,bL:result.leftover2.boxL,bW:result.leftover2.boxW,bH:result.leftover2.boxH},
    {col:0xa855f7,ox:result.leftover3.offX,oy:result.leftover3.offY,oz:result.leftover3.offZ,rnx:result.leftover3.nx,rny:result.leftover3.ny,rnz:result.leftover3.nz,bL:result.leftover3.boxL,bW:result.leftover3.boxW,bH:result.leftover3.boxH},
  ];
  const{cL,cW,cH}=result;
  useEffect(()=>{
    if(!mountRef.current) return;
    if(cleanRef.current){cleanRef.current();cleanRef.current=null;}
    const t=setTimeout(()=>{
      const el=mountRef.current;if(!el) return;const W=el.clientWidth||800,H=380;
      const scene=new THREE.Scene();scene.background=new THREE.Color(0xeef2f7);
      const camera=new THREE.PerspectiveCamera(50,W/H,0.01,100000);
      const renderer=new THREE.WebGLRenderer({antialias:true,preserveDrawingBuffer:true});
      renderer.setSize(W,H);renderer.setPixelRatio(Math.min(window.devicePixelRatio,2));el.appendChild(renderer.domElement);
      if(captureRef) captureRef.current=()=>{try{return renderer.domElement.toDataURL("image/png");}catch(e){return null;}};
      scene.add(new THREE.AmbientLight(0xffffff,0.85));
      const d=new THREE.DirectionalLight(0xffffff,0.7);d.position.set(5,8,5);scene.add(d);
      const d2=new THREE.DirectionalLight(0x8888ff,0.35);d2.position.set(-3,-2,-2);scene.add(d2);
      const tot=regs.reduce((s,r)=>s+((r.rnx||0)*(r.rny||0)*(r.rnz||0)),0)||1;
      regs.forEach(r=>{if(!r.rnx||!r.rny||!r.rnz||!r.bL||!r.bW||!r.bH) return;
        const total=r.rnx*r.rny*r.rnz;const cap=Math.max(1,Math.round(MAX_3D*(total/tot)));
        const stride=total>cap?total/cap:1;const pos=[];
        let idx=0,nextPick=0;
        for(let iz=0;iz<r.rnz;iz++)for(let iy=0;iy<r.rny;iy++)for(let ix=0;ix<r.rnx;ix++){
          // Always include boundary boxes so the full extent of the region is visible
          const isEdge=(ix===0||ix===r.rnx-1||iy===0||iy===r.rny-1||iz===0||iz===r.rnz-1);
          if(isEdge||idx>=nextPick){
            pos.push([(r.ox||0)+ix*r.bL+r.bL/2,(r.oz||0)+iz*r.bH+r.bH/2,(r.oy||0)+iy*r.bW+r.bW/2]);
            if(idx>=nextPick) nextPick+=stride;
          }
          idx++;
        }
        const mesh=new THREE.InstancedMesh(new THREE.BoxGeometry(r.bL*0.88,r.bH*0.88,r.bW*0.88),
          new THREE.MeshPhongMaterial({color:r.col,shininess:50,transparent:true,opacity:0.92}),pos.length);
        mesh.count=pos.length;const dummy=new THREE.Object3D();
        pos.forEach(([x,y,z],i)=>{dummy.position.set(x,y,z);dummy.updateMatrix();mesh.setMatrixAt(i,dummy.matrix);});
        mesh.instanceMatrix.needsUpdate=true;scene.add(mesh);});
      const cw=new THREE.LineSegments(new THREE.EdgesGeometry(new THREE.BoxGeometry(cL,cH,cW)),new THREE.LineBasicMaterial({color:0x1e293b}));
      cw.position.set(cL/2,cH/2,cW/2);scene.add(cw);
      const grid=new THREE.GridHelper(Math.max(cL,cW)*2.5,10,0xb0bec5,0xcfd8dc);grid.position.set(cL/2,-0.01,cW/2);scene.add(grid);
      const center=new THREE.Vector3(cL/2,cH/2,cW/2),diag=Math.sqrt(cL*cL+cW*cW+cH*cH)||10;
      let radius=diag*2.2,theta=Math.PI*0.38,phi=Math.PI*0.30;
      function updCam(){camera.position.set(center.x+radius*Math.sin(phi)*Math.sin(theta),center.y+radius*Math.cos(phi),center.z+radius*Math.sin(phi)*Math.cos(theta));camera.lookAt(center);}
      updCam();let drag=false,prev={x:0,y:0};const cv=renderer.domElement;cv.style.cursor="grab";cv.style.display="block";
      const onD=(e)=>{drag=true;prev={x:e.clientX,y:e.clientY};cv.style.cursor="grabbing";};
      const onM=(e)=>{if(!drag)return;theta-=(e.clientX-prev.x)*0.007;phi=Math.max(0.05,Math.min(1.5,phi+(e.clientY-prev.y)*0.007));prev={x:e.clientX,y:e.clientY};updCam();};
      const onU=()=>{drag=false;cv.style.cursor="grab";};
      const onW=(e)=>{e.preventDefault();radius=Math.max(diag*0.4,Math.min(diag*6,radius+e.deltaY*0.5));updCam();};
      const onRz=()=>{const nW=el.clientWidth||800;renderer.setSize(nW,H);camera.aspect=nW/H;camera.updateProjectionMatrix();};
      cv.addEventListener("mousedown",onD);window.addEventListener("mousemove",onM);window.addEventListener("mouseup",onU);
      cv.addEventListener("wheel",onW,{passive:false});window.addEventListener("resize",onRz);
      let animId;const loop=()=>{animId=requestAnimationFrame(loop);renderer.render(scene,camera);};loop();
      cleanRef.current=()=>{cancelAnimationFrame(animId);cv.removeEventListener("mousedown",onD);window.removeEventListener("mousemove",onM);
        window.removeEventListener("mouseup",onU);cv.removeEventListener("wheel",onW);window.removeEventListener("resize",onRz);
        if(el.contains(cv))el.removeChild(cv);renderer.dispose();};
    },100);
    return()=>{clearTimeout(t);if(cleanRef.current){cleanRef.current();cleanRef.current=null;}};
  },[result,regions3D]);
  return(<div ref={mountRef} style={{width:"100%",height:"380px",borderRadius:"10px",overflow:"hidden",border:"1px solid #e2e8f0",background:"#eef2f7"}}/>);}


export default ThreeViewer;
