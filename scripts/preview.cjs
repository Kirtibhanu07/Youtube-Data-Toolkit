const fs=require('node:fs'),http=require('node:http');
const sample=`<script>
document.getElementById('previewBanner').classList.remove('hidden');
let previewState={state:'EMPTY',message:'Start a collection or reload an existing job.'},success=()=>{},failure=()=>{};
const functions={toolkitStatus:()=>previewState,toolkitSelection:()=> 'https://www.youtube.com/watch?v=abcdefghijk\\nhttps://youtu.be/ABCDEFGHIJK',toolkitStart:()=>previewState={id:'preview',mode:document.getElementById('mode').value,state:'READY',phase:'METADATA',counts:{PENDING:100},total:100,apiAttempts:0,message:'Ready to collect. This preview uses sample data.'},toolkitRun:()=>{let done=Math.min(100,(previewState.counts.OK||0)+50);return previewState={...previewState,state:done===100?'COMPLETE':'READY',counts:{OK:done,PENDING:100-done},apiAttempts:previewState.apiAttempts+2,message:done===100?'Sample collection complete. No real YouTube data was fetched.':'Processing sample rows.'};},toolkitCancel:()=>previewState={...previewState,state:'CANCELLED',message:'Preview job cancelled.'},toolkitRetry:()=>previewState={...previewState,state:'READY',counts:{PENDING:100},message:'Ready to refresh sample rows.'}};
const runner=new Proxy({}, {get:(_,name)=>name==='withSuccessHandler'?(fn)=>(success=fn,runner):name==='withFailureHandler'?(fn)=>(failure=fn,runner):(...args)=>{const ok=success,fail=failure;setTimeout(()=>{try{ok(functions[name](...args));}catch(e){fail(e);}},350);}});
window.google={script:{run:runner}};
</script>`;
const html=fs.readFileSync('src/Sidebar.html','utf8').replace('<!--PREVIEW-->',sample);
http.createServer((req,res)=>{res.writeHead(200,{'Content-Type':'text/html; charset=utf-8'});res.end(html);}).listen(8504,'127.0.0.1',()=>console.log('Preview: http://127.0.0.1:8504 (sample data only)'));
