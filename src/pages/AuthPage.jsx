// ─── AUTH PAGE ────────────────────────────────────────────────────────────────
import { useState } from 'react';
import { useAuth } from '../contexts/AuthContext.jsx';
import DensiCubeLogo from '../components/DensiCubeLogo.jsx';

export default function AuthPage({ setPage }) {
  const { signIn, signUp, resetPassword } = useAuth();
  const [mode,setMode]=useState('login');
  const [name,setName]=useState('');
  const [email,setEmail]=useState('');
  const [password,setPassword]=useState('');
  const [loading,setLoading]=useState(false);
  const [error,setError]=useState('');
  const [success,setSuccess]=useState('');
  const reset=()=>{setError('');setSuccess('');};

  const handleSubmit=async()=>{
    reset();
    if(!email.trim()){setError('Please enter your email.');return;}
    if(mode!=='reset'&&!password){setError('Please enter your password.');return;}
    if(mode==='signup'&&!name.trim()){setError('Please enter your name.');return;}
    if(mode!=='reset'&&password.length<6){setError('Password must be at least 6 characters.');return;}
    setLoading(true);
    if(mode==='login'){
      const{error}=await signIn(email,password);
      if(error){setError(error.message);setLoading(false);return;}
      setPage('home');
    }else if(mode==='signup'){
      const{error}=await signUp(email,password,name);
      if(error){setError(error.message);setLoading(false);return;}
      setSuccess('Account created! Check your email to confirm, then log in.');
      setMode('login');
    }else{
      const{error}=await resetPassword(email);
      if(error){setError(error.message);setLoading(false);return;}
      setSuccess('Reset email sent! Check your inbox.');
    }
    setLoading(false);
  };

  const inp={width:'100%',padding:'11px 14px',border:'1px solid #e2e8f0',borderRadius:'8px',
    fontSize:'14px',fontFamily:'inherit',outline:'none',boxSizing:'border-box',color:'#1a2332'};
  const lbl={fontSize:'12px',fontWeight:'700',color:'#374151',textTransform:'uppercase',
    letterSpacing:'0.05em',display:'block',marginBottom:'5px'};

  return(
    <div style={{minHeight:'100vh',display:'flex',alignItems:'center',justifyContent:'center',
      background:'#f0fdf4',padding:'24px'}}>
      <div style={{width:'100%',maxWidth:'420px'}}>
        <div style={{textAlign:'center',marginBottom:'32px',cursor:'pointer'}} onClick={()=>setPage('home')}>
          <DensiCubeLogo size={48}/>
          <div style={{fontWeight:'800',fontSize:'22px',color:'#0f172a',marginTop:'8px'}}>DensiCube</div>
          <div style={{fontSize:'12px',color:'#94a3b8',letterSpacing:'0.1em',textTransform:'uppercase',marginTop:'2px'}}>Container Intelligence</div>
        </div>
        <div style={{background:'#fff',borderRadius:'16px',boxShadow:'0 4px 24px rgba(0,0,0,0.08)',padding:'32px'}}>
          <h2 style={{fontWeight:'800',fontSize:'20px',color:'#0f172a',margin:'0 0 4px'}}>
            {mode==='login'?'Welcome back':mode==='signup'?'Create your account':'Reset password'}
          </h2>
          <p style={{fontSize:'13px',color:'#6b7280',margin:'0 0 24px'}}>
            {mode==='login'?'Log in to your DensiCube account':
             mode==='signup'?'Free forever for single-container packing':
             "We'll send a reset link to your email"}
          </p>
          <div style={{display:'flex',flexDirection:'column',gap:'12px'}}>
            {mode==='signup'&&<div><label style={lbl}>Full Name</label>
              <input style={inp} type="text" value={name} onChange={e=>setName(e.target.value)} placeholder="Your name" autoFocus/></div>}
            <div><label style={lbl}>Email</label>
              <input style={inp} type="email" value={email} onChange={e=>setEmail(e.target.value)}
                placeholder="you@company.com" autoFocus={mode!=='signup'}
                onKeyDown={e=>e.key==='Enter'&&handleSubmit()}/></div>
            {mode!=='reset'&&<div><label style={lbl}>Password</label>
              <input style={inp} type="password" value={password} onChange={e=>setPassword(e.target.value)}
                placeholder={mode==='signup'?'Min. 6 characters':'••••••••'}
                onKeyDown={e=>e.key==='Enter'&&handleSubmit()}/></div>}
            {mode==='login'&&<div style={{textAlign:'right',marginTop:'-4px'}}>
              <button onClick={()=>{setMode('reset');reset();}}
                style={{background:'none',border:'none',color:'#be185d',fontSize:'12px',fontWeight:'600',cursor:'pointer',fontFamily:'inherit'}}>
                Forgot password?</button></div>}
          </div>
          {error&&<div style={{background:'#fff1f2',border:'1px solid #fecaca',borderRadius:'8px',
            padding:'10px 14px',fontSize:'13px',color:'#991b1b',marginTop:'14px'}}>⚠ {error}</div>}
          {success&&<div style={{background:'#f0fdf4',border:'1px solid #bbf7d0',borderRadius:'8px',
            padding:'10px 14px',fontSize:'13px',color:'#166534',marginTop:'14px'}}>✓ {success}</div>}
          <button onClick={handleSubmit} disabled={loading}
            style={{width:'100%',marginTop:'20px',padding:'12px',
              background:loading?'#e2e8f0':'linear-gradient(135deg,#be185d,#9d174d)',
              color:loading?'#9ca3af':'#fff',border:'none',borderRadius:'10px',
              fontWeight:'700',fontSize:'15px',cursor:loading?'not-allowed':'pointer',fontFamily:'inherit'}}>
            {loading?'⏳ Please wait...':mode==='login'?'Log In →':mode==='signup'?'Create Account →':'Send Reset Email →'}
          </button>
          <div style={{textAlign:'center',marginTop:'20px',fontSize:'13px',color:'#6b7280'}}>
            {mode==='login'&&<>Don't have an account?{' '}
              <button onClick={()=>{setMode('signup');reset();}}
                style={{background:'none',border:'none',color:'#be185d',fontWeight:'700',cursor:'pointer',fontFamily:'inherit',fontSize:'13px'}}>
                Sign up free</button></>}
            {mode==='signup'&&<>Already have an account?{' '}
              <button onClick={()=>{setMode('login');reset();}}
                style={{background:'none',border:'none',color:'#be185d',fontWeight:'700',cursor:'pointer',fontFamily:'inherit',fontSize:'13px'}}>
                Log in</button></>}
            {mode==='reset'&&<>Remember it?{' '}
              <button onClick={()=>{setMode('login');reset();}}
                style={{background:'none',border:'none',color:'#be185d',fontWeight:'700',cursor:'pointer',fontFamily:'inherit',fontSize:'13px'}}>
                Back to login</button></>}
          </div>
        </div>
        <div style={{textAlign:'center',marginTop:'20px'}}>
          <button onClick={()=>setPage('home')}
            style={{background:'none',border:'none',color:'#6b7280',fontSize:'13px',cursor:'pointer',fontFamily:'inherit'}}>
            ← Back to home</button>
        </div>
      </div>
    </div>
  );
}
