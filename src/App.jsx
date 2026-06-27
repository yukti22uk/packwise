// ─── DENSICUBE ROOT APP ────────────────────────────────────────────────────────
// Clean entry point — all logic lives in algorithms/, components/, tools/, pages/
import { useState, useEffect } from 'react';

// Pages
import Nav      from './pages/Nav.jsx';
import Footer   from './pages/Footer.jsx';
import HomePage from './pages/HomePage.jsx';
import ToolPage from './pages/ToolPage.jsx';
import PricingPage from './pages/PricingPage.jsx';
import AboutPage   from './pages/AboutPage.jsx';
import DensiCubeBot from './components/DensiCubeBot.jsx';

// ── Error Boundary ────────────────────────────────────────────────────────────
import { Component } from 'react';
class ErrorBoundary extends Component {
  constructor(props){ super(props); this.state={ hasError:false, error:null }; }
  static getDerivedStateFromError(error){ return{ hasError:true, error }; }
  componentDidCatch(error, info){ console.error('DensiCube error:', error, info); }
  render(){
    if(!this.state.hasError) return this.props.children;
    return(
      <div style={{padding:'48px 32px',textAlign:'center',maxWidth:'600px',margin:'0 auto'}}>
        <div style={{fontSize:'40px',marginBottom:'16px'}}>⚠️</div>
        <h2 style={{fontWeight:'800',color:'#111',marginBottom:'8px'}}>Something went wrong</h2>
        <p style={{color:'#6b7280',marginBottom:'24px'}}>
          An unexpected error occurred in this section.
          Your other tabs are unaffected.
        </p>
        <pre style={{background:'#f8fafc',borderRadius:'8px',padding:'12px',
          fontSize:'11px',color:'#ef4444',textAlign:'left',overflowX:'auto',
          marginBottom:'20px',maxHeight:'120px',overflowY:'auto'}}>
          {this.state.error?.message}
        </pre>
        <button onClick={()=>this.setState({hasError:false,error:null})}
          style={{padding:'10px 24px',background:'#be185d',color:'#fff',border:'none',
          borderRadius:'8px',fontWeight:'700',cursor:'pointer',fontSize:'14px'}}>
          Try Again
        </button>
      </div>
    );
  }
}

// ── Analytics (privacy-friendly, no cookies) ──────────────────────────────────
function trackEvent(name, props={}){
  // Uses Plausible if installed, otherwise console (dev mode)
  if(window.plausible){
    window.plausible(name, { props });
  }
}
export { trackEvent };

// ── Root App ──────────────────────────────────────────────────────────────────
export default function App(){
  const[page,setPage]=useState('home');
  const[isPro,setIsPro]=useState(false);
  const[modalOpen,setModalOpen]=useState(false);
  const[botTab,setBotTab]=useState(null);

  const handleBotNavigate=(tab)=>{
    setPage('tool');
    setBotTab(tab);
  };

  useEffect(()=>{
    try{ if(localStorage.getItem('pp_pro')==='true') setIsPro(true); }catch(e){}
  },[]);

  useEffect(()=>{
    window.scrollTo({top:0,behavior:'instant'});
    trackEvent('pageview', {page});
  },[page]);

  const logout=()=>{
    setIsPro(false);
    try{localStorage.removeItem('pp_pro');}catch(e){}
  };
  const openUpgrade=()=>{
    setPage('tool');
    setTimeout(()=>setModalOpen(true),100);
  };

  return(
    <div style={{fontFamily:"'Space Grotesk','Segoe UI',system-ui,-apple-system,sans-serif",
      background:'#f0fdf4',minHeight:'100vh'}}>
      <Nav page={page} setPage={setPage} isPro={isPro}
        onUpgrade={openUpgrade} onLogout={logout}/>
      <main>
        <ErrorBoundary key={page}>
          <div className="page-enter">
            {page==='home'    && <HomePage    setPage={setPage} onUpgrade={openUpgrade} onToolSelect={handleBotNavigate}/>}
            {page==='tool'    && <ToolPage    isPro={isPro} setIsPro={setIsPro}
                                   modalOpen={modalOpen} setModalOpen={setModalOpen}
                                   initialTab={botTab} onTabMounted={()=>setBotTab(null)}/>}
            {page==='pricing' && <PricingPage onUpgrade={openUpgrade} setPage={setPage}/>}
            {page==='about'   && <AboutPage   setPage={setPage}/>}
          </div>
        </ErrorBoundary>
      </main>
      <Footer setPage={setPage} onToolSelect={handleBotNavigate}/>
      <DensiCubeBot onNavigate={handleBotNavigate}/>
    </div>
  );
}
