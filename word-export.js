"use strict";
async function exportNativeWord(){
  if(!parsedData || exportingPDF) return;
  if(location.protocol !== 'http:' || location.hostname !== '127.0.0.1'){
    alert('完整 Word 导出需要本地服务。请关闭此页面，双击“启动排版工具.bat”，在自动打开的页面中重新导入书稿。'); return;
  }
  const controls=[...document.querySelectorAll('#toolbar button,#panel input,#panel select,#panel button,#emptyImport,#fileInput')];
  const previous=controls.map(el=>el.disabled),button=document.getElementById('exportWord');
  controls.forEach(el=>el.disabled=true);button.textContent='Word 正在排版…';
  const status=document.getElementById('stInfo');document.getElementById('statusBar').style.display='flex';
  status.textContent='正在调用本机 Word 排版并计算目录页码，请稍候…';
  try{
    readForm();
    const chapterArtwork=[];
    if(!String(P.chapterStyle).startsWith('版式一')){
      status.textContent='正在保留章节页视觉效果（章节页以整页图像嵌入，正文可编辑）…';
      if(renderTimer){clearTimeout(renderTimer);renderTimer=null;}
      await renderAll();await document.fonts.ready;
      if(document.getElementById('stInfo').textContent.startsWith('排版出错'))throw new Error('预览排版失败，请先修正书稿');
      pagesEl.style.setProperty('--zoom','1');
      for(const chapter of pagesEl.querySelectorAll('.chapter-page')){
        const sheet=chapter.closest('.pagedjs_sheet');if(!sheet)throw new Error('未找到章节页');
        const canvas=await htmlToImage.toCanvas(sheet,{width:sheet.offsetWidth,height:sheet.offsetHeight,pixelRatio:300/96,backgroundColor:getComputedStyle(sheet).backgroundColor,skipFonts:true,style:{margin:'0',transform:'none',zoom:'1',boxShadow:'none'}});
        chapterArtwork.push(canvas.toDataURL('image/png').split(',')[1]);canvas.width=canvas.height=1;
      }
      if(chapterArtwork.length!==parsedData.chapters.filter(ch=>ch.title!=null).length)throw new Error('章节页数量不完整，请检查预览');
    }
    const session=await fetch('/api/session');if(!session.ok)throw new Error('本地服务未就绪');
    const {token}=await session.json(),files=[];
    const encode=file=>new Promise((resolve,reject)=>{const reader=new FileReader();reader.onload=()=>resolve(String(reader.result).split(',')[1]);reader.onerror=()=>reject(new Error('文件读取失败：'+file.name));reader.readAsDataURL(file);});
    // Upload only the selected manuscript and images that it actually references.
    files.push({path:markdownRelPath,data:await encode(folderFiles.get(markdownRelPath))});
    const uploaded=new Set([markdownRelPath]);
    for(const path of collectImgPaths(parsedData)){
      const file=resolveImgPath(path);if(!file)throw new Error('找不到插图：'+path+'。请将图片放入导入的书稿文件夹。');
      const rel=(file.webkitRelativePath||file.name).replace(/\\/g,'/');
      if(!uploaded.has(rel)){files.push({path:rel,data:await encode(file)});uploaded.add(rel);}
    }
    const response=await fetch('/api/export-word',{method:'POST',headers:{'Content-Type':'application/json','X-Export-Token':token},body:JSON.stringify({source:markdownRelPath,name:BOOK_NAME,params:P,files,chapterArtwork})});
    if(!response.ok)throw new Error(await response.text());
    const blob=await response.blob(),url=URL.createObjectURL(blob),link=document.createElement('a');
    link.href=url;link.download=(BOOK_NAME||'书稿').replace(/[\\/:*?"<>|]/g,'_')+'.docx';document.body.appendChild(link);link.click();link.remove();
    setTimeout(()=>URL.revokeObjectURL(url),60000);status.textContent='完整 Word 已生成，目录页码已由本机 Word 计算。';
  }catch(error){status.textContent='Word 导出失败：'+error.message;alert(status.textContent);}
  finally{applyZoom();controls.forEach((el,i)=>el.disabled=previous[i]);button.textContent='导出完整 Word';}
}
document.getElementById('exportWord').addEventListener('click',exportNativeWord);
