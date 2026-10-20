const fs=require('node:fs');const {harness,item}=require('../test/harness.cjs');
const cases=[50,500,1000].map(n=>{
 const ids=Array.from({length:n},(_,i)=>String(i).padStart(11,'0'));
 const h=harness({Videos:(_,o)=>({items:o.id.split(',').map(id=>item(id))}),Channels:()=>({items:[]})});
 const j=h.ctx.toolkitStart({mode:'videos',source:ids.join('\n')});let state=j;
 for(let i=0;i<n/50+3&&!state.state.startsWith('COMPLETE');i++)state=h.ctx.toolkitRun(j.id);
 const videos=h.calls.filter(c=>c.resource==='Videos').length,channels=h.calls.filter(c=>c.resource==='Channels').length;
 return {inputRows:n,successfulRows:state.counts.OK,videoListAttempts:videos,channelListAttempts:channels,totalListAttempts:h.calls.length,
  perVideoBaselineVideoListAttempts:n,videoListCallReductionPercent:100*(1-videos/n),state:state.state};
});
const out={kind:'Offline deterministic service mocks, not a live YouTube benchmark',cases};
fs.mkdirSync('docs',{recursive:true});fs.writeFileSync('docs/benchmark.json',JSON.stringify(out,null,2)+'\n');console.log(JSON.stringify(out,null,2));
