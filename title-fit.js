"use strict";
(function(root){
  function calculate(measure,base,minimum,width,enabled=true){
    if(!enabled)return {size:base,fits:true,enabled:false};
    const floor=Math.min(base,minimum),limit=Math.max(0,width)*.96;
    let size=base;
    while(size>floor && measure(size)>limit)size=Math.max(floor,Math.floor((size-.01)*2)/2);
    return {size,fits:measure(size)<=limit,enabled:true};
  }
  const api={calculate};
  if(typeof module!=="undefined" && module.exports)module.exports=api;
  else root.TitleFit=api;
})(globalThis);
