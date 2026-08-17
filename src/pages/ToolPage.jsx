// ─── TOOL PAGE ───────────────────────────────────────────────────────────────
import { useState, useRef, useEffect } from 'react';
import UpgradeModal from '../components/UpgradeModal.jsx';
import ProGate from '../components/ProGate.jsx';
import BoxPackingTool from '../tools/BoxPackingTool.jsx';
import ShipmentPlanner from '../tools/ShipmentPlanner.jsx';
import MultiSKUTool from '../tools/MultiSKUTool.jsx';
import ContainerSkuTool from '../tools/BulkSKUTool.jsx';
import SKUGrouperTool from '../tools/SKUGrouperTool.jsx';
import OrderAnalyserTool     from '../tools/OrderAnalyserTool.jsx';
import WarehouseDesignerTool from '../tools/WarehouseDesignerTool.jsx';
function ToolPage({isPro,setIsPro,modalOpen,setModalOpen,initialTab,onTabMounted}){
  const[tab,setTab]=useState(initialTab||"box");
  const[multiSKUPreset,setMultiSKUPreset]=useState(null);

  useEffect(()=>{
    if(initialTab){setTab(initialTab);if(onTabMounted)onTabMounted();}
  },[initialTab]);
  const unlock=()=>{setIsPro(true);try{localStorage.setItem("pp_pro","true");}catch(e){}setTimeout(()=>setModalOpen(false),1200);};
  const tabs=[
    ["box","📦 Single SKU Calculator",false],
    ["multisku","🗃️ Multi-SKU Planner",true],
    ["shipment","🚚 Shipment Planner",true],
    ["sku","🗃️ Bulk SKU Calculator",false],
    ["grouper","🔀 SKU Grouper",false],
    ["analyser","📊 Order Analyser",true],
    ["warehouse","🏭 WH Designer",true],
  ];
  const handleSendToMultiSKU=(rows)=>{setMultiSKUPreset(rows);setTab("multisku");};
  return(
    <div>
      <UpgradeModal open={modalOpen} onClose={()=>setModalOpen(false)} onUnlock={unlock}/>
      {/* Tool sub-nav */}
      <div style={{background:"linear-gradient(135deg,#0f172a 0%,#0d2b1a 100%)",padding:"20px 32px 0"}}>
        <div style={{maxWidth:"1200px",margin:"0 auto"}}>
          <p style={{color:"#94a3b8",fontSize:"12px",margin:"0 0 14px"}}>
            Single SKU · Multi-SKU · Shipment Planner · Bulk SKU · SKU Grouper · Order Analyser
          </p>
          <div style={{display:"flex",gap:"4px",flexWrap:"wrap"}}>
            {tabs.map(([id,label,pro])=>(
              <button key={id} onClick={()=>setTab(id)} style={{padding:"9px 18px",border:"none",cursor:"pointer",
                fontSize:"13px",fontWeight:"600",borderRadius:"8px 8px 0 0",
                background:tab===id?"#f0f4f8":"rgba(255,255,255,0.08)",
                color:tab===id?"#1a2332":"#cbd5e1",transition:"all 0.15s",position:"relative"}}>
                {label}{pro&&!isPro&&<span style={{marginLeft:"5px",fontSize:"10px",background:"#fbbf24",color:"#78350f",padding:"1px 5px",borderRadius:"99px",fontWeight:"700"}}>PRO</span>}
              </button>
            ))}
          </div>
        </div>
      </div>
      <div style={{background:"#f0fdf4",minHeight:"60vh",padding:"24px 32px 0"}}>
        <div style={{maxWidth:"1200px",margin:"0 auto"}}>
          {tab==="box"&&<BoxPackingTool/>}
          {tab==="multisku"&&(isPro?<MultiSKUTool preset={multiSKUPreset} onPresetUsed={()=>setMultiSKUPreset(null)}/>:<ProGate feature="Multi-SKU Planner" onUpgrade={()=>setModalOpen(true)}/>)}
          {tab==="shipment"&&(isPro?<ShipmentPlanner/>:<ProGate feature="Shipment Planner" onUpgrade={()=>setModalOpen(true)}/>)}
          {tab==="sku"&&<ContainerSkuTool isPro={isPro} onUpgrade={()=>setModalOpen(true)}/>}
          {tab==="grouper"&&<SKUGrouperTool onSendToMultiSKU={handleSendToMultiSKU}/>}
          {tab==="analyser"&&(isPro?<OrderAnalyserTool/>:<ProGate feature="Order Analyser" onUpgrade={()=>setModalOpen(true)}/>)}
          {tab==="warehouse"&&(isPro?<WarehouseDesignerTool/>:<ProGate feature="Warehouse Designer" onUpgrade={()=>setModalOpen(true)}/>)}
        </div>
      </div>
    </div>
  );
}


export default ToolPage;
