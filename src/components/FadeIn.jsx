// ─── SCROLL FADE-IN ──────────────────────────────────────────────────────────
import { useState, useRef, useEffect } from 'react';
function FadeIn({children,className="",style={},stagger=false}){
  const ref=useRef(null);
  useEffect(()=>{
    const el=ref.current;if(!el) return;
    const obs=new IntersectionObserver(([e])=>{
      if(e.isIntersecting){el.classList.add("in");obs.disconnect();}
    },{threshold:0.1,rootMargin:"0px 0px -40px 0px"});
    obs.observe(el);return()=>obs.disconnect();
  },[]);
  return <div ref={ref} className={`${stagger?"stagger":"fade-up"} ${className}`} style={style}>{children}</div>;
}

// Animated count-up number (triggers on scroll into view)

export default FadeIn;
