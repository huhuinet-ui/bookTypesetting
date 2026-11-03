"use strict";
(()=>{
  const dialog=document.getElementById('wordImportDialog'),status=document.getElementById('wordImportStatus'),report=document.getElementById('wordImportReport'),download=document.getElementById('downloadMarkdown');
  let downloadURL=null;
  document.getElementById('importWord').addEventListener('click',()=>{
    if(location.protocol!=='http:'||location.hostname!=='127.0.0.1'){
      alert('Word 转 Markdown 需要本地服务，请双击“启动排版工具.bat”后使用。');return;
    }
    dialog.showModal();
  });
  document.getElementById('closeWordImport').addEventListener('click',()=>dialog.close());
  const decode=data=>Uint8Array.from(atob(data),c=>c.charCodeAt(0));
  document.getElementById('convertWord').addEventListener('click',async()=>{
    const file=document.getElementById('wordSource').files[0];
    if(!file||!file.name.toLowerCase().endsWith('.docx')){status.textContent='请先选择一个 .docx 文件。';return;}
    if(parsedData&&!confirm('转换成功后将替换当前载入的书稿，排版参数会保留。是否继续？'))return;
    const button=document.getElementById('convertWord');button.disabled=true;
    status.textContent='正在转换并提取图片，请稍候…';report.textContent='';download.hidden=true;
    if(downloadURL){URL.revokeObjectURL(downloadURL);downloadURL=null;}
    try{
      if(file.size>150*1024*1024)throw new Error('文件过大，请使用 150 MB 以内的 Word 文档。');
      const session=await fetch('/api/session');if(!session.ok)throw new Error('本地服务未就绪');
      const {token}=await session.json();
      const data=await new Promise((resolve,reject)=>{const reader=new FileReader();reader.onload=()=>resolve(String(reader.result).split(',')[1]);reader.onerror=()=>reject(new Error('读取文件失败'));reader.readAsDataURL(file);});
      const response=await fetch('/api/convert-word',{method:'POST',headers:{'Content-Type':'application/json','X-Export-Token':token},body:JSON.stringify({name:file.name,data,bracketSubtitle:document.getElementById('bracketSubtitle').checked})});
      if(!response.ok)throw new Error(await response.text());
      const result=await response.json(),files=result.files.map(item=>{
        const f=new File([decode(item.data)],item.path.split('/').pop(),{type:item.path.endsWith('.md')?'text/markdown':'application/octet-stream'});
        Object.defineProperty(f,'webkitRelativePath',{value:'转换书稿/'+item.path});return f;
      });
      downloadURL=URL.createObjectURL(new Blob([decode(result.archive)],{type:'application/zip'}));
      download.href=downloadURL;download.download=file.name.replace(/\.docx$/i,'')+'_Markdown.zip';download.hidden=false;
      report.textContent=result.report;loadFolder(files);
      status.textContent='转换完成，已载入书稿。可关闭窗口查看预览，或下载 ZIP 保存 Markdown 与图片。';
    }catch(error){status.textContent='转换失败：'+error.message;}
    finally{button.disabled=false;}
  });
})();
