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
    let stage='连接本地服务';
    try{
      if(file.size>1024*1024*1024)throw new Error('文件过大，请使用 1 GB 以内的 Word 文档。');
      const session=await fetch('/api/session');if(!session.ok)throw new Error('本地服务未就绪');
      const {token,wordBinaryUpload,wordFileResults}=await session.json();
      if(!wordBinaryUpload || !wordFileResults)throw new Error('本地服务仍是旧版本。请关闭原启动窗口，重新双击“启动排版工具.bat”，使用新打开的页面重试。');
      stage='上传或转换 Word';
      status.textContent='正在上传并转换 Word，请稍候…';
      const response=await fetch('/api/convert-word',{method:'POST',headers:{'Content-Type':'application/octet-stream','X-Export-Token':token,'X-Word-Name':encodeURIComponent(file.name),'X-Word-Results':'files','X-Bracket-Subtitle':String(document.getElementById('bracketSubtitle').checked)},body:file});
      if(!response.ok)throw new Error(await response.text());
      stage='接收转换结果';
      status.textContent='转换完成，正在接收文字和图片…';
      const result=await response.json(),files=[];
      download.href=result.archiveUrl;download.download=file.name.replace(/\.docx$/i,'')+'_Markdown.zip';download.hidden=false;
      for(const [index,item] of result.files.entries()){
        status.textContent=`正在接收文字和图片 ${index+1}/${result.files.length}…`;
        const asset=await fetch(item.url);
        if(!asset.ok)throw new Error('转换文件读取失败，请重新转换：'+item.path);
        const f=new File([await asset.blob()],item.path.split('/').pop(),{type:item.path.endsWith('.md')?'text/markdown':'application/octet-stream'});
        Object.defineProperty(f,'webkitRelativePath',{value:'转换书稿/'+item.path});files.push(f);
      }
      report.textContent=result.report;loadFolder(files);
      status.textContent='转换完成，已载入书稿。可关闭窗口查看预览，或下载 ZIP 保存 Markdown 与图片。';
    }catch(error){
      status.textContent='转换失败：'+(error instanceof TypeError && /fetch|network|load failed/i.test(error.message)?stage+'时连接中断。请检查本地启动窗口是否仍在运行，并查看窗口中的错误信息；重启后请使用新打开的页面，旧页面的端口可能已失效。':error.message);
    }
    finally{button.disabled=false;}
  });
})();
