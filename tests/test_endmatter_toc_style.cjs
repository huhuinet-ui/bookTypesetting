const fs=require('node:fs'),path=require('node:path'),vm=require('node:vm'),assert=require('node:assert/strict');
const source=fs.readFileSync(path.join(__dirname,'..','排版工具.html'),'utf8');
let style=1;
const ctx=vm.createContext({P:{pad:3,chapterTocSize:13},esc:v=>v,isIllustratedToc:()=>style===2,isReferenceToc:()=>style===1,isVerticalToc:()=>style===3,renderTocEntry:()=>'<div class="toc-entry">Closing</div>'});
vm.runInContext(source.slice(source.indexOf('function endmatterEntries('),source.indexOf('function renderVerticalToc(')),ctx);
const data={chapters:[{},{}],afterword:'Text',afterwordTitle:'Closing',appendices:[{title:'Appendix'}]};
let html=ctx.renderAfterwordToc(data);
assert(html.includes('toc-section-title'));
assert(!html.includes('toc-reference-row'));
assert(html.includes('data-toc-ref="afterword"'));
assert(html.includes('class="toc-afterword '));
style=2;html=ctx.renderAfterwordToc(data);assert(html.includes('toc-illustrated-chapter'));
style=3;html=ctx.renderAfterwordToc(data);assert(html.includes('toc-vertical-label">C\nl\no\ns\ni\nn\ng'));
assert(html.includes('toc-vertical-chapter-newpage'));
assert(html.includes('data-toc-ref="appendix0"'));
assert.equal(ctx.renderAfterwordToc({afterword:null,chapters:[],appendices:[]}), '');
data.appendices[0].articles=[{title:'Supplement article',author:''}];
for(style=1;style<4;style++){
  const articleHtml=ctx.renderAfterwordToc(data);
  assert(articleHtml.includes('Supplement article'));
  assert(articleHtml.includes('data-toc-ref="appendix0_article0"'));
}
data.endmatter_order=[{kind:'appendix',index:0},{kind:'afterword'}];
for(style=0;style<4;style++){
  const ordered=ctx.endmatterEntries(data);
  assert.equal(ordered[0].key,'appendix0');assert.equal(ordered[1].key,'afterword');
  if(style){const rendered=ctx.renderAfterwordToc(data);assert(rendered.indexOf('data-toc-ref="appendix0"')<rendered.indexOf('data-toc-ref="afterword"'));}
}
console.log('Endmatter entries use chapter layout in every TOC style');
