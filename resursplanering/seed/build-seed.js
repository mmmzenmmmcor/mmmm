// Bygger startdata (från befintlig Excel-planering v.44–46 2026) som JSON.
const people = [
  ['jorgen-tufvesson','Jörgen Tufvesson','Sydbyggarna',''],
  ['pontus-nilsson','Pontus Nilsson','Sydbyggarna',''],
  ['andre-nordstrand','Andre Nordstrand','Sydbyggarna',''],
  ['bjorn-olofsson','Björn Olofsson','Sydbyggarna',''],
  ['erik-jonsson','Erik Jönsson','Sydbyggarna','Lärling'],
  ['wictor-holmstrom','Wictor Holmström','Sydbyggarna',''],
  ['samir-azizi','Samir Azizi','Sydbyggarna','Skolepåg'],
  ['marcin-olszewski','Marcin Wojciech Olszewski','Sydbyggarna',''],
  ['iulian-duma-sic','Iulian Duma-Sic','Dezignhus',''],
  ['gavril-dereleanu','Gavril Dereleanu','Dezignhus',''],
  ['gabriel-voidoc','Gabriel Voidoc','Dezignhus',''],
  ['elmiers-pontus','Pontus','Elmiers Bygg',''],
  ['elmiers-teo','Teo','Elmiers Bygg',''],
];
const projects = [['1039','#E8894A'],['1044','#B48CE8'],['1045','#8DB8E0'],['1046','#EBCB6B'],['1048','#B9BDB8'],['10087','#A9CC9A']];
const P = n => ({p:'p'+n}), N = (n,note) => ({p:'p'+n, n:note}), L = {s:'ledig'};
const all = v => ({1:v,2:v,3:v,4:v,5:v});
const wedFri = v => ({3:v,4:v,5:v});
const weeks = {
  '2026-W44': {
    'jorgen-tufvesson': all(P(1045)),
    'pontus-nilsson': {1:P(1039),2:P(1039),3:N(1039,'Ledig på eftermiddagen'),4:P(1039),5:P(1039)},
    'andre-nordstrand': all(P(1044)), 'bjorn-olofsson': all(L), 'erik-jonsson': all(P(1044)),
    'wictor-holmstrom': all(L), 'samir-azizi': all(P(1039)), 'marcin-olszewski': all(P(1039)),
    'iulian-duma-sic': all(P(1039)), 'gabriel-voidoc': all(P(1039)),
    'elmiers-pontus': all(P(1039)), 'elmiers-teo': all(P(1039)),
  },
  '2026-W45': {
    'jorgen-tufvesson': all(P(1045)), 'pontus-nilsson': all(P(1039)),
    'andre-nordstrand': all(P(10087)), 'bjorn-olofsson': all(P(1039)), 'erik-jonsson': all(P(10087)),
    'wictor-holmstrom': wedFri(P(1039)), 'samir-azizi': wedFri(P(1039)), 'marcin-olszewski': all(P(1039)),
    'iulian-duma-sic': all(P(1044)), 'gabriel-voidoc': all(P(1044)),
    'elmiers-pontus': all(P(1039)), 'elmiers-teo': all(P(1039)),
  },
  '2026-W46': {
    'jorgen-tufvesson': all(P(1045)),
    'pontus-nilsson': {1:P(1039),2:P(1039),3:N(1039,'Ledig på eftermiddagen'),4:P(1039),5:P(1039)},
    'andre-nordstrand': {1:P(10087),2:P(1044),3:P(1044),4:P(1044),5:P(1044)},
    'bjorn-olofsson': all(P(1039)),
    'erik-jonsson': {1:P(10087),2:P(1044),3:P(1044),4:P(1044),5:P(1044)},
    'wictor-holmstrom': wedFri(P(1039)), 'samir-azizi': wedFri(P(1039)), 'marcin-olszewski': all(P(1039)),
    'iulian-duma-sic': {1:P(1044),2:N(1048,'Etablering'),3:P(1044),4:N(1046,'Demontering'),5:P(1044)},
    'gabriel-voidoc': {1:P(1044),2:N(1048,'Etablering'),3:P(1044),4:N(1046,'Demontering'),5:P(1044)},
    'elmiers-pontus': all(P(1039)), 'elmiers-teo': all(P(1039)),
  },
};
const ops = [];
people.forEach(([id,name,company,tag],i) => ops.push({action:'set', collection:'people', doc_id:id, data:{name,company,tag,order:i,active:true}}));
projects.forEach(([num,color]) => ops.push({action:'set', collection:'projects', doc_id:'p'+num, data:{number:num,name:'',client:'',color,active:true}}));
for (const [k,cells] of Object.entries(weeks)) ops.push({action:'set', collection:'weeks', doc_id:k, data:{key:k, cells}});
require('fs').writeFileSync(__dirname+'/seed.json', JSON.stringify(ops,null,1));
console.log(ops.length+' dokument');
