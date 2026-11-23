"use strict";
let projectMarkdownName="";
(()=>{
  const input=document.getElementById("projectInput"),status=document.getElementById("projectStatus");
  document.getElementById("saveProject").addEventListener("click",()=>{
    readForm();
    const data={format:BookProjectFormat.FORMAT,version:BookProjectFormat.VERSION,savedAt:new Date().toISOString(),source:{markdownName:markdownRelPath.split("/").pop()||projectMarkdownName},settings:{...P},pdfQuality:document.getElementById("pdfQuality").value};
    try{BookProjectFormat.normalize(data,GROUPS);}catch(error){status.textContent="保存失败："+error.message;return;}
    const url=URL.createObjectURL(new Blob([JSON.stringify(data,null,2)],{type:"application/json;charset=utf-8"}));
    const link=document.createElement("a");link.href=url;link.download=(BOOK_NAME||projectMarkdownName.replace(/\.md$/i,"")||"书稿").replace(/[\\/:*?"<>|]/g,"_")+"_排版工程.json";
    document.body.appendChild(link);link.click();link.remove();setTimeout(()=>URL.revokeObjectURL(url),10000);
    status.textContent="已下载排版工程。下次导入此 JSON，再选择原书稿文件夹即可恢复。";
  });
  document.getElementById("openProject").addEventListener("click",()=>{input.value="";input.click();});
  input.addEventListener("change",async()=>{
    const file=input.files[0];if(!file)return;
    try{
      if(file.size>2*1024*1024)throw new Error("工程文件过大，请选择 2 MB 以内的 JSON 文件。");
      const project=BookProjectFormat.normalize(JSON.parse((await file.text()).replace(/^\uFEFF/,"")),GROUPS);
      // Validate the complete document before changing any current settings.
      Object.assign(P,project.settings);
      projectMarkdownName=project.source.markdownName;
      document.getElementById("pdfQuality").value=project.pdfQuality;
      buildForm();filterSettings();updateSpineEstimate();
      const current=markdownRelPath.split("/").pop();
      status.textContent="已恢复工程设置。"+(parsedData?(projectMarkdownName && current!==projectMarkdownName?"当前书稿与工程记录的文件名不同，请选择原书稿："+projectMarkdownName:"已应用到当前书稿。"):"请导入书稿文件夹"+(projectMarkdownName?"（"+projectMarkdownName+"）":"")+"。");
      if(parsedData)scheduleRender();
    }catch(error){status.textContent="导入失败："+error.message;}
  });
})();
