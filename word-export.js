"use strict";
let cachedWord=null;
function downloadWordBlob(blob){
  const url=URL.createObjectURL(blob),link=document.createElement('a');
  link.href=url;link.download=(BOOK_NAME||'书稿').replace(/[\\/:*?"<>|]/g,'_')+'.docx';document.body.appendChild(link);link.click();link.remove();
  setTimeout(()=>URL.revokeObjectURL(url),60000);
}
async function exportNativeWord(){
  if(!parsedData || exportingPDF) return;
  if(location.protocol !== 'http:' || location.hostname !== '127.0.0.1'){
    alert('完整 Word 导出需要本地服务。请关闭此页面，双击“启动排版工具.bat”，在自动打开的页面中重新导入书稿。'); return;
  }
  const controls=[...document.querySelectorAll('#toolbar button,#panel input,#panel select,#panel textarea,#panel button,#emptyImport,#fileInput')];
  const previous=controls.map(el=>el.disabled),button=document.getElementById('exportWord');
  controls.forEach(el=>el.disabled=true);button.textContent='Word 正在排版…';
  const status=document.getElementById('stInfo');document.getElementById('statusBar').style.display='flex';
  status.textContent='正在调用本机 Word 排版并计算目录页码，请稍候…';
  try{
    readForm();
    const sourceFile=folderFiles.get(markdownRelPath), paramsKey=JSON.stringify(P);
    if(cachedWord && cachedWord.book===parsedData && cachedWord.source===sourceFile && cachedWord.key===paramsKey){
      downloadWordBlob(cachedWord.blob);status.textContent='Word 已下载（复用本次书稿的排版结果）。';return;
    }
    const chapterArtwork=[];
    // Rasterize only the small ornament once; all chapter text stays native Word text.
    if(!String(P.chapterStyle).startsWith('版式一')){
      status.textContent='正在准备章节装饰，文字将以可编辑格式导出…';
      const svg=isChap5()?chap5PlaneSVG():isChap4()?chap4RoseSVG():chap2OrnamentSVG();
      const url=URL.createObjectURL(new Blob([svg],{type:'image/svg+xml'}));
      try{
        const image=new Image();image.src=url;await image.decode();
        const canvas=document.createElement('canvas');
        canvas.width=isChap2()?800:400;canvas.height=isChap2()?1100:500;
        canvas.getContext('2d').drawImage(image,0,0,canvas.width,canvas.height);
        chapterArtwork.push(canvas.toDataURL('image/png').split(',')[1]);
        canvas.width=canvas.height=1;
      }finally{URL.revokeObjectURL(url);}
    }
    const session=await fetch('/api/session');if(!session.ok)throw new Error('本地服务未就绪');
    const {token}=await session.json(),files=[];
    const encode=file=>new Promise((resolve,reject)=>{const reader=new FileReader();reader.onload=()=>resolve(String(reader.result).split(',')[1]);reader.onerror=()=>reject(new Error('文件读取失败：'+(file.name||'书稿或图片')+'。请重新导入书稿文件夹后再导出。'));reader.readAsDataURL(file);});
    // Upload only the selected manuscript and images that it actually references.
    files.push({path:markdownRelPath,data:await encode(markdownSnapshot || folderFiles.get(markdownRelPath))});
    const uploaded=new Set([markdownRelPath]);
    for(const path of collectImgPaths(parsedData)){
      const file=resolveImgPath(path);if(!file)throw new Error('找不到插图：'+path+'。请将图片放入导入的书稿文件夹。');
      const rel=(file.webkitRelativePath||file.name).replace(/\\/g,'/');
      if(!uploaded.has(rel)){files.push({path:rel,data:await encode(imgInfo.get(path)?.sourceBlob || file)});uploaded.add(rel);}
    }
    const response=await fetch('/api/export-word',{method:'POST',headers:{'Content-Type':'application/json','X-Export-Token':token},body:JSON.stringify({source:markdownRelPath,name:BOOK_NAME,params:P,files,chapterArtwork})});
    if(!response.ok)throw new Error(await response.text());
    const blob=await response.blob();
    cachedWord={book:parsedData,source:sourceFile,key:paramsKey,blob};
    downloadWordBlob(blob);status.textContent='完整 Word 已生成，目录页码已由本机 Word 计算。';
  }catch(error){status.textContent='Word 导出失败：'+error.message;alert(status.textContent);}
  finally{applyZoom();controls.forEach((el,i)=>el.disabled=previous[i]);button.textContent='导出完整 Word';}
}
document.getElementById('exportWord').addEventListener('click',exportNativeWord);
