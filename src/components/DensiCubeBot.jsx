// ─── DENSICUBE CHATBOT ─────────────────────────────────────────────────────────
// Floating chat widget that helps users find the right tool.
import { useState, useRef, useEffect } from 'react';

const TOOLS = [
  { id:'box',      label:'📦 Single SKU Calculator', tab:'box' },
  { id:'multisku', label:'🗃️ Multi-SKU Planner',     tab:'multisku' },
  { id:'shipment', label:'🚚 Shipment Planner',       tab:'shipment' },
  { id:'sku',      label:'🗃️ Bulk SKU Calculator',   tab:'sku'      },
  { id:'grouper',  label:'🔀 SKU Grouper',            tab:'grouper'  },
  { id:'analyser', label:'📊 Order Analyser',         tab:'analyser' },
];

const PROBLEMS = [
  {
    id: 'p1',
    question: 'How many boxes fit in my truck / container?',
    tool: 'box',
    answer: 'The Single SKU Calculator tells you the exact number of boxes that fit in any container or vehicle — with a 3D view and 2D loading diagram.',
  },
  {
    id: 'p2',
    question: 'I have multiple box sizes — how do I pack them together?',
    tool: 'multisku',
    answer: 'The Multi-SKU Planner packs up to 8 different box sizes into one container simultaneously and shows you a colour-coded 3D loading plan.',
  },
  {
    id: 'p3',
    question: 'How many trucks do I need for my full order?',
    tool: 'shipment',
    answer: 'The Shipment Planner takes your total order quantity and tells you exactly how many containers are needed, with a per-container manifest and PDF loading plan.',
  },
  {
    id: 'p4',
    question: 'I have a large catalogue — how many of each SKU fits?',
    tool: 'sku',
    answer: 'The Bulk SKU Calculator processes hundreds of SKUs from an Excel file at once and tells you the maximum quantity per SKU for your chosen container.',
  },
  {
    id: 'p5',
    question: 'I have 10,000+ SKUs — how do I group them for shipping?',
    tool: 'grouper',
    answer: 'The SKU Grouper clusters your entire catalogue into 2–8 representative size groups using AI, then sends them straight to the Multi-SKU Planner.',
  },
  {
    id: 'p6',
    question: 'Which products should I prioritise for shipping?',
    tool: 'analyser',
    answer: 'The Order Analyser gives you ABC analysis (by volume), FMS classification (by order frequency), and an ABC-FMS matrix — showing exactly which SKUs to prioritise.',
  },
  {
    id: 'p7',
    question: 'I want to compare freight cost across vehicle types',
    tool: 'shipment',
    answer: 'The Shipment Planner has a built-in cost comparison table — enter freight cost per vehicle type and it instantly calculates cost per unit for each option.',
  },
  {
    id: 'p8',
    question: 'My order data has errors — how do I find them?',
    tool: 'analyser',
    answer: 'The Order Analyser flags all anomalies in your master SKU and order data — missing dimensions, zero quantities, missing dates, duplicate SKUs, and more.',
  },
  {
    id: 'p9',
    question: 'How do I plan shipment for multiple SKUs across many trucks?',
    tool: 'shipment',
    answer: 'Use the Shipment Planner in Multi-SKU mode — it plans how multiple box sizes fill each container and gives you a per-container manifest for every truck.',
  },
  {
    id: 'p10',
    question: 'I need a loading plan PDF for my warehouse team',
    tool: 'box',
    answer: 'Both the Single SKU Calculator and Shipment Planner export a branded PDF loading plan with dimensions, orientation, stacking instructions, and a 3D diagram.',
  },
];

const WELCOME = {
  id: 'welcome',
  from: 'bot',
  text: "Hi! I'm the DensiCube assistant 👋\n\nWhat problem are you trying to solve? Select one below and I'll point you to the right tool.",
  showOptions: true,
};

export default function DensiCubeBot({ onNavigate }) {
  const [open,     setOpen]    = useState(false);
  const [messages, setMessages]= useState([WELCOME]);
  const [phase,    setPhase]   = useState('problems');

  // Auto-open on first visit, then remember preference
  useEffect(() => {
    const seen = localStorage.getItem('pw_bot_seen');
    if (!seen) {
      setTimeout(() => setOpen(true), 1200);
      localStorage.setItem('pw_bot_seen', '1');
    }
  }, []); // 'problems' | 'done'
  const bottomRef = useRef(null);

  useEffect(() => {
    if (open) bottomRef.current?.scrollIntoView({ behavior: 'smooth' });
  }, [messages, open]);

  const reset = () => { setMessages([WELCOME]); setPhase('problems'); };

  const handleProblem = (prob) => {
    const tool = TOOLS.find(t => t.id === prob.tool);

    setMessages(prev => [
      ...prev,
      // User selection bubble
      { id: Date.now()+'u', from: 'user', text: prob.question },
      // Bot answer
      { id: Date.now()+'b', from: 'bot', text: prob.answer,
        tool, showTool: true },
    ]);
    setPhase('done');
  };

  const handleGoToTool = (tab) => {
    onNavigate(tab);
    setOpen(false);
  };

  return (
    <>
      {/* Floating button */}
      <button onClick={() => setOpen(o => !o)}
        style={{ position:'fixed', bottom:'24px', right:'24px', zIndex:1000,
          width:'56px', height:'56px', borderRadius:'50%',
          background:'linear-gradient(135deg,#be185d,#9d174d)',
          border:'none', cursor:'pointer', boxShadow:'0 4px 20px rgba(190,24,93,0.45)',
          display:'flex', alignItems:'center', justifyContent:'center',
          fontSize:'24px', transition:'transform 0.2s',
          transform: open ? 'scale(0.9)' : 'scale(1)' }}
        aria-label="Open DensiCube assistant">
        {open ? '✕' : '💬'}
      </button>

      {/* Chat window */}
      {open && (
        <div style={{ position:'fixed', bottom:'92px', right:'24px', zIndex:999,
          width:'360px', maxWidth:'calc(100vw - 32px)',
          background:'#fff', borderRadius:'16px',
          boxShadow:'0 8px 40px rgba(0,0,0,0.18)', overflow:'hidden',
          display:'flex', flexDirection:'column', maxHeight:'560px' }}>

          {/* Header */}
          <div style={{ background:'linear-gradient(135deg,#be185d,#9d174d)',
            padding:'14px 18px', display:'flex', justifyContent:'space-between',
            alignItems:'center', flexShrink:0 }}>
            <div style={{ display:'flex', alignItems:'center', gap:'10px' }}>
              <div style={{ width:'36px', height:'36px', borderRadius:'50%',
                background:'rgba(255,255,255,0.2)', display:'flex',
                alignItems:'center', justifyContent:'center', fontSize:'18px' }}>🤖</div>
              <div>
                <div style={{ fontWeight:'700', color:'#fff', fontSize:'14px' }}>DensiCube Assistant</div>
                <div style={{ fontSize:'11px', color:'rgba(255,255,255,0.75)' }}>Find the right tool instantly</div>
              </div>
            </div>
            <button onClick={reset}
              style={{ background:'rgba(255,255,255,0.15)', border:'none', color:'#fff',
                borderRadius:'8px', padding:'4px 10px', cursor:'pointer',
                fontSize:'11px', fontWeight:'600', fontFamily:'inherit' }}>
              Restart
            </button>
          </div>

          {/* Messages */}
          <div style={{ flex:1, overflowY:'auto', padding:'16px', display:'flex',
            flexDirection:'column', gap:'12px', background:'#f8fafc' }}>
            {messages.map((msg) => (
              <div key={msg.id}>
                {/* Bubble */}
                <div style={{ display:'flex', justifyContent: msg.from==='user' ? 'flex-end' : 'flex-start' }}>
                  {msg.from==='bot' && (
                    <div style={{ width:'28px', height:'28px', borderRadius:'50%', flexShrink:0,
                      background:'#be185d', display:'flex', alignItems:'center',
                      justifyContent:'center', fontSize:'14px', marginRight:'8px',
                      alignSelf:'flex-end' }}>🤖</div>
                  )}
                  <div style={{ maxWidth:'80%', padding:'10px 14px', borderRadius:
                      msg.from==='user' ? '16px 16px 4px 16px' : '16px 16px 16px 4px',
                    background: msg.from==='user' ? '#be185d' : '#fff',
                    color: msg.from==='user' ? '#fff' : '#1a2332',
                    fontSize:'13px', lineHeight:'1.5', boxShadow:'0 1px 3px rgba(0,0,0,0.08)',
                    whiteSpace:'pre-line' }}>
                    {msg.text}
                  </div>
                </div>

                {/* Tool link card */}
                {msg.showTool && msg.tool && (
                  <div style={{ marginLeft:'36px', marginTop:'8px' }}>
                    <div style={{ background:'#fff', border:'1px solid #e2e8f0',
                      borderRadius:'12px', padding:'12px 14px',
                      boxShadow:'0 1px 4px rgba(0,0,0,0.06)' }}>
                      <div style={{ fontSize:'11px', fontWeight:'700', color:'#9ca3af',
                        textTransform:'uppercase', letterSpacing:'0.06em', marginBottom:'8px' }}>
                        Recommended Tool
                      </div>
                      <button onClick={() => handleGoToTool(msg.tool.tab)}
                        style={{ width:'100%', padding:'10px 14px',
                          background:'linear-gradient(135deg,#be185d,#9d174d)',
                          color:'#fff', border:'none', borderRadius:'8px',
                          fontWeight:'700', fontSize:'13px', cursor:'pointer',
                          fontFamily:'inherit', textAlign:'left',
                          display:'flex', alignItems:'center', justifyContent:'space-between' }}>
                        <span>{msg.tool.label}</span>
                        <span style={{ fontSize:'16px' }}>→</span>
                      </button>
                    </div>
                  </div>
                )}

                {/* Problem options */}
                {msg.showOptions && (
                  <div style={{ marginLeft:'36px', marginTop:'10px',
                    display:'flex', flexDirection:'column', gap:'6px' }}>
                    {PROBLEMS.map(prob => (
                      <button key={prob.id} onClick={() => handleProblem(prob)}
                        style={{ padding:'9px 12px', background:'#fff',
                          border:'1px solid #e2e8f0', borderRadius:'10px',
                          fontSize:'12px', fontWeight:'600', color:'#374151',
                          cursor:'pointer', textAlign:'left', fontFamily:'inherit',
                          transition:'all 0.15s', lineHeight:'1.4' }}
                        onMouseOver={e => { e.currentTarget.style.borderColor='#be185d'; e.currentTarget.style.background='#fdf2f8'; }}
                        onMouseOut={e => { e.currentTarget.style.borderColor='#e2e8f0'; e.currentTarget.style.background='#fff'; }}>
                        {prob.question}
                      </button>
                    ))}
                  </div>
                )}
              </div>
            ))}

            {/* After answer — ask another */}
            {phase === 'done' && (
              <div style={{ display:'flex', justifyContent:'center', marginTop:'4px' }}>
                <button onClick={reset}
                  style={{ padding:'7px 16px', background:'#f1f5f9',
                    border:'1px solid #e2e8f0', borderRadius:'99px',
                    fontSize:'12px', fontWeight:'600', color:'#6b7280',
                    cursor:'pointer', fontFamily:'inherit' }}>
                  ← Ask another question
                </button>
              </div>
            )}

            <div ref={bottomRef}/>
          </div>
        </div>
      )}
    </>
  );
}
