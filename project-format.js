"use strict";
// Shared by the browser and the format regression tests.
(function(root){
  const FORMAT="book-typesetting-project", VERSION=1;
  function normalize(data,groups){
    if(!data || data.format!==FORMAT)throw new Error("这不是图书排版工程文件。");
    if(data.version!==VERSION)throw new Error("不支持此工程版本，请使用兼容版本的排版工具。");
    if(!data.settings || typeof data.settings!=="object" || Array.isArray(data.settings))throw new Error("工程缺少排版设置。");
    const settings={};
    for(const group of groups)for(const item of group.items){
      let value=Object.hasOwn(data.settings,item.k)?data.settings[item.k]:item.d;
      if(item.k==='chapterTextColor' && !Object.hasOwn(data.settings,item.k)){
        const legacy=/纸飞机|指纹树/.test(data.settings.chapterStyle||'')?'chapter5TitleColor':'chapter4TitleColor';
        value=data.settings[legacy]||item.d;
      }
      let valid=false;
      if(item.t==="num")valid=typeof value==="number" && Number.isFinite(value) && (item.min==null||value>=item.min) && (item.max==null||value<=item.max);
      else if(item.t==="bool")valid=typeof value==="boolean";
      else if(item.t==="select")valid=item.opts.includes(value);
      else if(item.t==="color")valid=typeof value==="string" && /^#[\da-f]{6}$/i.test(value);
      else valid=typeof value==="string";
      if(!valid)throw new Error("工程参数无效："+item.l);
      settings[item.k]=value;
    }
    const source=data.source && typeof data.source==="object"?data.source:{};
    return {settings,source:{markdownName:typeof source.markdownName==="string"?source.markdownName:""},pdfQuality:data.pdfQuality==="print"?"print":"fast"};
  }
  const api={FORMAT,VERSION,normalize};
  if(typeof module!=="undefined" && module.exports)module.exports=api;
  else root.BookProjectFormat=api;
})(globalThis);
