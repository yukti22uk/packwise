// ─── ANIMATED COUNT-UP ───────────────────────────────────────────────────────
import { useState, useRef, useEffect } from 'react';
function CountUp({value,suffix="",prefix="",duration=1600}){
  const[n,setN]=useState(0);const ref=useRef(null);const done=useRef(false);
  useEffect(()=>{
    const el=ref.current;if(!el) return;
    const obs=new IntersectionObserver(([e])=>{
      if(e.isIntersecting&&!done.current){
        done.current=true;const start=Date.now();
        const tick=()=>{
          const p=Math.min((Date.now()-start)/duration,1);
          const eased=1-Math.pow(1-p,3);
          setN(Math.round(value*eased));
          if(p<1) requestAnimationFrame(tick);
        };requestAnimationFrame(tick);obs.disconnect();}
    },{threshold:0.5});
    obs.observe(el);return()=>obs.disconnect();
  },[value,duration]);
  return <span ref={ref}>{prefix}{n}{suffix}</span>;
}

// WhatsApp share button

export default CountUp;
