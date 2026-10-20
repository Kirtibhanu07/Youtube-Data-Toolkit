const fs=require('node:fs'),vm=require('node:vm'),path=require('node:path');
function harness(api={}){
 const properties=new Map(),sheets=new Map();let nextId=1,locked=false,selection=[['https://youtu.be/abcdefghijk']];
 class Range{
  constructor(sheet,row,col,n=1,m=1){Object.assign(this,{sheet,row,col,n,m});}
  getValues(){return Array.from({length:this.n},(_,i)=>Array.from({length:this.m},(_,j)=>(this.sheet.cells[this.row+i-1]||[])[this.col+j-1]??''));}
  getDisplayValues(){return this.getValues().map(r=>r.map(String));}
  setValues(rows){if(rows.length!==this.n||rows.some(r=>r.length!==this.m))throw Error('Range dimensions mismatch');for(let i=0;i<this.n;i++){let row=this.sheet.cells[this.row+i-1]||[];for(let j=0;j<this.m;j++)row[this.col+j-1]=rows[i][j];this.sheet.cells[this.row+i-1]=row;}return this;}
  setBackground(){return this;}setFontColor(){return this;}setFontWeight(){return this;}setNote(){return this;}
 }
 class Sheet{
  constructor(name){this.id=nextId++;this.name=name;this.cells=[];this.maxRows=1000;this.maxCols=26;}
  getRange(...a){return new Range(this,...a);}getSheetId(){return this.id;}getLastRow(){return this.cells.length;}getMaxRows(){return this.maxRows;}getMaxColumns(){return this.maxCols;}
  insertRowsAfter(_,n){this.maxRows+=n;}insertColumnsAfter(_,n){this.maxCols+=n;}setFrozenRows(){}setColumnWidths(){}setColumnWidth(){}
 }
 const source=new Sheet('Source');source.cells=[['untouched input','do not replace']];sheets.set(source.id,source);
 const ss={getId:()=> 'sheet-1',getUrl:()=> 'https://docs.google.com/spreadsheets/d/sheet-1/edit',insertSheet(name){let s=new Sheet(name);sheets.set(s.id,s);return s;},getSheetById:id=>sheets.get(id)};
 let calls=[],sleeps=[];const youtube={};
 for(const resource of ['Videos','Channels','Playlists','PlaylistItems'])youtube[resource]={list(part,options){calls.push({resource,part,options});if(api[resource])return api[resource](part,options,calls);return {items:[]};}};
 let uid=0;
 const ctx=vm.createContext({console,Date,Math,Set,Array,Object,JSON,Number,String,Error,isFinite,decodeURIComponent,
  SpreadsheetApp:{getActiveSpreadsheet:()=>ss,flush(){},getActiveRange:()=>({getNumColumns:()=>1,getNumRows:()=>selection.length,getDisplayValues:()=>selection})},
  PropertiesService:{getUserProperties:()=>({getProperty:k=>properties.get(k)||null,setProperty:(k,v)=>{if(Buffer.byteLength(v)>9000)throw Error('Property exceeds limit');properties.set(k,v);}})},
  LockService:{getUserLock:()=>({tryLock(){if(locked)return false;locked=true;return true;},releaseLock(){locked=false;}})},
  Utilities:{getUuid:()=>String(++uid).padStart(32,'0'),sleep:n=>sleeps.push(n),formatDate:d=>d.toISOString().replace('T',' ').slice(0,19)},Session:{getScriptTimeZone:()=> 'Asia/Kolkata'},YouTube:youtube});
 for(const n of ['Core.js','Code.js'])vm.runInContext(fs.readFileSync(path.join(__dirname,'../src',n),'utf8'),ctx,{filename:n});
 return {ctx,core:ctx.YTCore,calls,sleeps,sheets,source,properties,ss,setSelection:x=>selection=x};
}
const item=(id='abcdefghijk',extra={})=>({id,snippet:{title:'Example video',channelId:'UC'+'a'.repeat(22),channelTitle:'Example channel',publishedAt:'2025-01-31T23:59:59Z',description:'Description',liveBroadcastContent:'none'},statistics:{viewCount:'0',likeCount:'15'},contentDetails:{duration:'PT1H2M3S'},...extra});
module.exports={harness,item};
