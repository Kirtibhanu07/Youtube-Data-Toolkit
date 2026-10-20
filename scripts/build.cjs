const fs=require('node:fs');const path=require('node:path');
const root=path.resolve(__dirname,'..');fs.mkdirSync(path.join(root,'dist'),{recursive:true});
fs.writeFileSync(path.join(root,'dist/Code.gs'),['Core.js','Code.js'].map(n=>fs.readFileSync(path.join(root,'src',n),'utf8')).join('\n'));
for(const name of ['Sidebar.html','appsscript.json'])fs.copyFileSync(path.join(root,'src',name),path.join(root,'dist',name));
console.log('Built dist/Code.gs, Sidebar.html and appsscript.json');
