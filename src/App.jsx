// ─── DENSICUBE ROOT APP ───────────────────────────────────────────────────────
import { useState, useEffect } from 'react';
import { Component } from 'react';
import { AuthProvider, useAuth } from './contexts/AuthContext.jsx';
import Nav          from './pages/Nav.jsx';
import Footer       from './pages/Footer.jsx';
import HomePage     from './pages/HomePage.jsx';
import ToolPage     from './pages/ToolPage.jsx';
import PricingPage  from './pages/PricingPage.jsx';
import AboutPage    from './pages/AboutPage.jsx';
import AuthPage     from './pages/AuthPage.jsx';
import AccountPage  from './pages/AccountPage.jsx';
import DensiCubeBot from './components/DensiCubeBot.jsx';

// ── Error Boundary ─────────────────────────────────────────────────────────────
class ErrorBoundary extends Component {
  constructor(props){ super(props); this.state={hasError:false,error:null}; }
  static getDerivedStateFromError(e){ return{hasError:true,error:e}; }
  componentDidCatch(e,i){ console.error('DensiCube error:',e,i); }
  render(){
    if(!this.state.hasError) return this.props.children;
    return(
      <div style={{padding:'48px 32px',textAlign:'center',maxWidth:'600px',margin:'0 auto'}}>
        <div style={{fontSize:'40px',marginBottom:'16px'}}>⚠️</div>
        <h2 style={{fontWeight:'800',color:'#111',marginBottom:'8px'}}>Something went wrong</h2>
        <p style={{color:'#6b7280',marginBottom:'24px'}}>An unexpected error occurred. Your other tabs are unaffected.</p>
        <pre style={{background:'#f8fafc',borderRadius:'8px',padding:'12px',fontSize:'11px',
          color:'#ef4444',textAlign:'left',overflowX:'auto',marginBottom:'20px',maxHeight:'120px',overflowY:'auto'}}>
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

function trackEvent(name,props={}){ if(window.plausible) window.plausible(name,{props}); }
export{trackEvent};

// ── Inner App ──────────────────────────────────────────────────────────────────
function AppInner(){
  const{user,isPro,loading,signOut}=useAuth();
  const[page,setPage]=useState('home');
  const[modalOpen,setModalOpen]=useState(false);
  const[botTab,setBotTab]=useState(null);

  useEffect(()=>{ window.scrollTo({top:0,behavior:'instant'}); trackEvent('pageview',{page}); },[page]);

  const handleBotNavigate=(tab)=>{ setPage('tool'); setBotTab(tab); };

  const openUpgrade=()=>{
    if(!user){ setPage('auth'); return; }
    setPage('tool');
    setTimeout(()=>setModalOpen(true),100);
  };

  const handleLogout=async()=>{ await signOut(); setPage('home'); };

  if(loading) return(
    <div style={{minHeight:'100vh',display:'flex',alignItems:'center',justifyContent:'center',background:'#f0fdf4'}}>
      <div style={{textAlign:'center'}}>
        <div style={{fontSize:'32px',marginBottom:'12px'}}>⏳</div>
        <div style={{color:'#6b7280',fontSize:'14px'}}>Loading DensiCube...</div>
      </div>
    </div>
  );

  // Auth + Account pages — no Nav/Footer wrapper needed
  if(page==='auth')    return <AuthPage    setPage={setPage}/>;
  if(page==='account') return <AccountPage setPage={setPage}/>;

  return(
    <div style={{fontFamily:"'Space Grotesk','Segoe UI',system-ui,-apple-system,sans-serif",
      background:'#f0fdf4',minHeight:'100vh'}}>
      <Nav page={page} setPage={setPage} isPro={isPro}
        user={user} onLogin={()=>setPage('auth')}
        onUpgrade={openUpgrade} onLogout={handleLogout}
        onAccount={()=>setPage('account')}/>
      <main>
        <ErrorBoundary key={page}>
          <div className="page-enter">
            {page==='home'    &&<HomePage    setPage={setPage} onUpgrade={openUpgrade} onToolSelect={handleBotNavigate}/>}
            {page==='tool'    &&<ToolPage    isPro={isPro} setIsPro={()=>{}}
                                  modalOpen={modalOpen} setModalOpen={setModalOpen}
                                  initialTab={botTab} onTabMounted={()=>setBotTab(null)}/>}
            {page==='pricing' &&<PricingPage onUpgrade={openUpgrade} setPage={setPage}/>}
            {page==='about'   &&<AboutPage   setPage={setPage}/>}
          </div>
        </ErrorBoundary>
      </main>
      <Footer setPage={setPage} onToolSelect={handleBotNavigate}/>
      <DensiCubeBot onNavigate={handleBotNavigate}/>
    </div>
  );
}

export default function App(){
  return(
    <AuthProvider>
      <AppInner/>
    </AuthProvider>
  );
}
