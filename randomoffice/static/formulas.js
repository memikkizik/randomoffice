/* Small formula parser: arithmetic, cell references and numeric aggregates. */
function columnLabel(index) {
  let result = '';
  for (let n=index+1;n>0;n=Math.floor((n-1)/26)) result=String.fromCharCode(65+(n-1)%26)+result;
  return result;
}

function formulaValue(value, visiting = new Set()) {
  if (typeof value !== 'string' || !value.startsWith('=')) {
    if (value === '' || value === null || value === undefined) return null;
    const number = Number(value);
    return Number.isFinite(number) ? number : null;
  }
  if (visiting.size > 80 || value.length > 4000) throw new Error('#LIMIT!');
  const expression=value.slice(1).toUpperCase().replace(/\s/g,'');
  const tokens=expression.match(/(?:\d+(?:\.\d*)?|\.\d+)|[A-Z]+[1-9]\d*|[A-ZÇĞİÖŞÜ]+|[()+*/,:;^-]/g)||[];
  if(tokens.join('')!==expression) throw new Error('#VALUE!');
  const rows=excelModel.sheets[currentSheetIndex].rows;
  let position=0;
  const scalar=value=>{if(Array.isArray(value)) throw new Error('#VALUE!');return value??0;};
  const address=ref=>{
    const match=/^([A-Z]+)([1-9]\d*)$/.exec(ref||'');
    if(!match) throw new Error('#REF!');
    return [Number(match[2])-1,[...match[1]].reduce((n,c)=>n*26+c.charCodeAt(0)-64,0)-1];
  };
  const readCell=(r,c)=>{
    const key=`${r}:${c}`;
    if(visiting.has(key)) throw new Error('#CYCLE!');
    visiting.add(key);
    try{return formulaValue(rows[r]?.[c],visiting);}finally{visiting.delete(key);}
  };
  const expect=token=>{if(tokens[position++]!==token) throw new Error('#VALUE!');};
  function primary(){
    const token=tokens[position++];
    if(token==='+'||token==='-') return (token==='-'?-1:1)*scalar(primary());
    if(token==='('){const result=sum();expect(')');return result;}
    if(/^(?:\d|\.)/.test(token||'')) return Number(token);
    if(/^[A-Z]+[1-9]\d*$/.test(token||'')){
      const [r,c]=address(token);
      if(tokens[position]!==':') return readCell(r,c);
      position++;
      const [endR,endC]=address(tokens[position++]);
      if((Math.abs(endR-r)+1)*(Math.abs(endC-c)+1)>200000) throw new Error('#LIMIT!');
      const values=[];
      for(let row=Math.min(r,endR);row<=Math.max(r,endR);row++){
        for(let col=Math.min(c,endC);col<=Math.max(c,endC);col++) values.push(readCell(row,col));
      }
      return values;
    }
    const aliases={TOPLA:'SUM',ORTALAMA:'AVERAGE',MİN:'MIN',MAK:'MAX',SAY:'COUNT'};
    const name=aliases[token]||token;
    if(!['SUM','AVERAGE','MIN','MAX','COUNT'].includes(name)) throw new Error('#NAME?');
    expect('(');
    const values=[];
    if(tokens[position]!==')'){
      do{values.push(...[sum()].flat());}while([',',';'].includes(tokens[position])&&position++);
    }
    expect(')');
    const numbers=values.filter(value=>value!==null);
    const total=numbers.reduce((a,b)=>a+b,0);
    if(name==='SUM') return total;
    if(name==='COUNT') return numbers.length;
    if(name==='AVERAGE'){if(!numbers.length) throw new Error('#DIV/0!');return total/numbers.length;}
    return numbers.length?numbers.reduce((a,b)=>name==='MIN'?Math.min(a,b):Math.max(a,b)):0;
  }
  function product(){
    let value=primary();
    while(['*','/'].includes(tokens[position])){
      const operator=tokens[position++],right=scalar(primary());
      if(operator==='/'&&right===0) throw new Error('#DIV/0!');
      value=operator==='*'?scalar(value)*right:scalar(value)/right;
    }
    return value;
  }
  function sum(){
    let value=product();
    while(['+','-'].includes(tokens[position])){
      const operator=tokens[position++],right=scalar(product());
      value=operator==='+'?scalar(value)+right:scalar(value)-right;
    }
    return value;
  }
  const result=scalar(sum());
  if(position!==tokens.length||!Number.isFinite(result)) throw new Error('#VALUE!');
  return result;
}

function evaluateFormulaDisplay(value){
  if(typeof value!=='string'||!value.startsWith('=')) return value;
  try{return formulaValue(value).toLocaleString('tr-TR',{maximumFractionDigits:10});}
  catch(error){return error.message.startsWith('#')?error.message:'#VALUE!';}
}

function displayExcelCell(value,row,column){
  const format=excelModel.sheets[currentSheetIndex].number_formats?.[`${row}:${column}`];
  if(format&&value!==''&&value!==null&&value!==undefined){
    try{
      const number=formulaValue(value);
      if(number!==null&&format.includes('%')) return (number*100).toLocaleString('tr-TR',{maximumFractionDigits:2})+'%';
      if(number!==null&&format.includes('₺')) return number.toLocaleString('tr-TR',{style:'currency',currency:'TRY'});
    }catch(error){}
  }
  return evaluateFormulaDisplay(value);
}

function setExcelNumberFormat(format){
  const {r,c}=activeCellPos;
  const sheet=excelModel.sheets[currentSheetIndex];
  const cell=document.querySelector(`#excel-sheet-grid td[data-row="${r}"][data-col="${c}"]`);
  if(cell===document.activeElement) cell.blur();
  sheet.number_formats||={};
  sheet.number_formats[`${r}:${c}`]=format;
  if(cell) cell.innerText=displayExcelCell(sheet.rows[r]?.[c],r,c);
  persistState();
}
