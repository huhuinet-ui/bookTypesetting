"use strict";
const {test}=require("node:test"),assert=require("node:assert/strict"),fs=require("node:fs"),vm=require("node:vm");
const format=require("../project-format.js");
const html=fs.readFileSync(require("node:path").join(__dirname,"../排版工具.html"),"utf8");
const definitions=html.slice(html.indexOf("const FONT_SUGGEST"),html.indexOf("/* ==================== 本机主题色历史"));
const {groups,defaults}=vm.runInNewContext(definitions+";({groups:GROUPS,defaults:P})");
const make=settings=>({format:format.FORMAT,version:format.VERSION,settings,source:{markdownName:"当妈12年.md"},pdfQuality:"print"});
test("all current settings round trip, including custom text and fonts",()=>{
  const settings={...defaults,bodySize:12,bookColor:"#123456",titlePageText:'一封信\n{{文字}} <文字> "',fontBody:"自定义字体"};
  // Use the real schema, including every newly added control.
  const project=format.normalize(JSON.parse(JSON.stringify(make(settings))),groups);
  for(const group of groups)for(const item of group.items)assert.equal(project.settings[item.k],settings[item.k]);
  assert.equal(project.source.markdownName,"当妈12年.md");assert.equal(project.pdfQuality,"print");
});
test("older projects fill missing controls with defaults and ignore unknown keys",()=>{
  const project=format.normalize(make({bodySize:11,unknownSetting:"ignored"}),groups);
  assert.equal(project.settings.bodySize,11);assert.equal(project.settings.pageW,defaults.pageW);
  assert.equal(Object.hasOwn(project.settings,"unknownSetting"),false);
});
test("invalid file, future version and invalid settings fail atomically",()=>{
  for(const data of [null,{}, {...make({}),version:2},make([]),make({bodySize:"12"}),make({pageW:-1}),make({bookColor:'"><script>'}),make({titlePageOn:"false"})])assert.throws(()=>format.normalize(data,groups));
  assert.equal(defaults.bodySize,groups.flatMap(g=>g.items).find(i=>i.k==="bodySize").d);
});
test("all inline scripts compile",()=>{
  for(const match of html.matchAll(/<script\b[^>]*>([\s\S]*?)<\/script>/gi))new vm.Script(match[1]);
});
