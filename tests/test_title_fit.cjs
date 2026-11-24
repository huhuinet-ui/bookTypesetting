const {test}=require('node:test'),assert=require('node:assert/strict');
const {calculate}=require('../title-fit.js');
test('short titles keep their configured font size',()=>{
  assert.deepEqual(calculate(size=>size*4,14,10,100),{size:14,fits:true,enabled:true});
});
test('long titles shrink individually in half-point steps with Word safety space',()=>{
  const fit=calculate(size=>size*20,14,10,250);
  assert.equal(fit.size,12);assert.equal(fit.fits,true);
  assert.ok((fit.size+.5)*20>250*.96);
});
test('titles below capacity at minimum are flagged without endless shrinking',()=>{
  assert.deepEqual(calculate(size=>size*80,21,12,180),{size:12,fits:false,enabled:true});
});
test('disabled adaptation and minimum larger than configured size',()=>{
  assert.equal(calculate(size=>size*80,14,10,100,false).size,14);
  assert.equal(calculate(size=>size*80,8,12,100).size,8);
});
