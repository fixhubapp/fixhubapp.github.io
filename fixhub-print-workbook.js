/* Local file import and print-format export for FixHub. */
window.FixHubWorkbook = (() => {
  'use strict';
  const xml = 'application/xml';
  const mainNS = 'http://schemas.openxmlformats.org/spreadsheetml/2006/main';
  const relNS = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';
  const children = (n, name) => Array.from(n.children).filter(el => el.localName === name);
  const first = (n, name) => children(n, name)[0];
  const make = (n, name) => n.ownerDocument.createElementNS(n.namespaceURI || mainNS, name);
  const set = (node, attrs) => { for (const [name, value] of Object.entries(attrs)) node.setAttribute(name, value); return node; };
  function insert(parent, name, order) {
    let node = first(parent, name);
    if (!node) {
      node = make(parent, name);
      const index = order.indexOf(name);
      const next = Array.from(parent.children).find(el => order.indexOf(el.localName) > index);
      parent.insertBefore(node, next || null);
    }
    return node;
  }
  const sheetOrder = 'sheetPr dimension sheetViews sheetFormatPr cols sheetData sheetCalcPr sheetProtection protectedRanges scenarios autoFilter sortState dataConsolidate customSheetViews mergeCells phoneticPr conditionalFormatting dataValidations hyperlinks printOptions pageMargins pageSetup headerFooter rowBreaks colBreaks customProperties cellWatches ignoredErrors smartTags drawing legacyDrawing legacyDrawingHF picture oleObjects controls webPublishItems tableParts extLst'.split(' ');
  const bookOrder = 'fileVersion fileSharing workbookPr workbookProtection bookViews sheets functionGroups externalReferences definedNames calcPr oleSize customWorkbookViews pivotCaches smartTagPr smartTagTypes webPublishing fileRecoveryPr webPublishObjects extLst'.split(' ');
  function csvRows(text) {
    text = text.replace(/^\uFEFF/, '');
    if (!text.length) throw Error('This CSV is empty.');
    // Count separators in the first logical record, ignoring quoted fields.
    const counts = {',':0, ';':0, '\t':0};
    let quoted = false;
    for (let i=0; i<text.length; i++) {
      const c=text[i];
      if(c==='"'){if(quoted && text[i+1]==='"')i++;else quoted=!quoted;}
      else if(!quoted){if(c==='\r'||c==='\n')break;if(c in counts)counts[c]++;}
    }
    const sep = Object.keys(counts).sort((a,b)=>counts[b]-counts[a])[0];
    const rows=[]; let row=[], value='', inside=false, ended=false;
    for(let i=0;i<text.length;i++){
      const c=text[i];ended=false;
      if(inside){
        if(c==='"'){if(text[i+1]==='"'){value+='"';i++;}else inside=false;}
        else value+=c;
      }else if(c==='"' && !value.length)inside=true;
      else if(c===sep){row.push(value);value='';}
      else if(c==='\r'||c==='\n'){if(c==='\r' && text[i+1]==='\n')i++;row.push(value);rows.push(row);row=[];value='';ended=true;}
      else value+=c;
    }
    if(inside)throw Error('This CSV has an unclosed quoted field.');
    if(!ended){row.push(value);rows.push(row);}
    return rows;
  }
  function fromRows(rows, name) {
    const book=XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(book, XLSX.utils.aoa_to_sheet(rows), name);
    return book;
  }
  async function read(file) {
    const bytes=new Uint8Array(await file.arrayBuffer());
    let book, original=null;
    if(/\.csv$/i.test(file.name)){
      const enc=bytes[0]===255 && bytes[1]===254?'utf-16le':bytes[0]===254 && bytes[1]===255?'utf-16be':'utf-8';
      let text;
      try{text=new TextDecoder(enc,{fatal:true}).decode(bytes);}
      catch{throw Error('Please save this CSV as UTF-8 or upload an XLSX copy so the characters can be preserved.');}
      book=fromRows(csvRows(text), 'Sheet1');
    }else if(/\.xlsx$/i.test(file.name)){
      if(bytes[0]!==80 || bytes[1]!==75)throw Error('This file is not a readable XLSX workbook.');
      book=XLSX.read(bytes,{type:'array',cellStyles:true,cellFormula:true,sheetStubs:true,cellNF:true});
      original=bytes;
    }else throw Error('Choose an Excel (.xlsx) or CSV (.csv) file.');
    return {name:file.name,book,original};
  }
  function model(book, name) {
    const sheet=book.Sheets[name], hiddenRows=sheet['!rows']||[], hiddenCols=sheet['!cols']||[];
    const cells=new Map();let loR=Infinity,hiR=-1,loC=Infinity,hiC=-1,uncached=0;
    for(const addr of Object.keys(sheet)){
      if(addr[0]==='!')continue;
      const cell=sheet[addr];
      if((cell.v==null||cell.v==='') && !cell.f)continue;
      const {r,c}=XLSX.utils.decode_cell(addr);
      if(hiddenRows[r]?.hidden||hiddenCols[c]?.hidden)continue;
      let text;
      if(cell.f && (cell.v==null || cell.t==='z')){text='='+cell.f;uncached++;}
      else text=cell.w ?? XLSX.utils.format_cell(cell);
      if(!cells.has(r))cells.set(r,new Map());
      cells.get(r).set(c,String(text));loR=Math.min(loR,r);hiR=Math.max(hiR,r);loC=Math.min(loC,c);hiC=Math.max(hiC,c);
    }
    const rows=[],cols=[];
    if(hiR>=0){
      for(let r=loR;r<=hiR;r++)if(!hiddenRows[r]?.hidden)rows.push(r);
      for(let c=loC;c<=hiC;c++)if(!hiddenCols[c]?.hidden)cols.push(c);
    }
    return {name,rows,cols,cells,uncached};
  }
  const partPath=(parent,target)=>decodeURIComponent(new URL(target,'https://workbook.invalid/'+parent).pathname.slice(1));
  const relsPath=part=>part.replace(/(^|\/)([^/]+)$/,'$1_rels/$2.rels');
  async function doc(zip,path){
    const file=zip.file(path);
    if(!file)throw Error('The workbook is missing its '+path+' part.');
    const d=new DOMParser().parseFromString(await file.async('string'),xml);
    if(d.getElementsByTagName('parsererror').length)throw Error('The workbook contains unreadable XML.');
    return d;
  }
  const save=(zip,path,d)=>zip.file(path,new XMLSerializer().serializeToString(d).replace(/\r/g,'&#13;'));
  function updateStyles(styles){
    const root=styles.documentElement,xfs=first(root,'cellXfs'),fonts=first(root,'fonts');
    const oldXfs=children(xfs,'xf'),oldFonts=children(fonts,'font'),mapped=new Map(),fontMap=new Map();
    const baseFont=oldFonts[Number(oldXfs[0]?.getAttribute('fontId'))]||oldFonts[0];
    const normalName=first(baseFont,'name')?.getAttribute('val')||'Calibri';
    const normalSize=Number(first(baseFont,'sz')?.getAttribute('val'))||11;
    const ctx=document.createElement('canvas').getContext('2d');
    ctx.font=(normalSize*96/72)+'px "'+normalName.replace(/"/g,'')+'"';
    const digitWidth=Math.max(1,Math.round(Math.max(...Array.from('0123456789',c=>ctx.measureText(c).width))));
    function getStyle(id,heading){
      const key=id+':'+heading;
      if(mapped.has(key))return mapped.get(key);
      const xf=(oldXfs[id]||oldXfs[0]).cloneNode(true);
      const fontId=Number(xf.getAttribute('fontId'))||0, fk=fontId+':'+heading;
      if(!fontMap.has(fk)){
        const font=(oldFonts[fontId]||baseFont).cloneNode(true);
        const field=(n,v)=>{let e=first(font,n);if(!e){e=make(font,n);font.append(e);}e.setAttribute('val',v);};
        first(font,'scheme')?.remove();field('name','Arial');field('sz','10');
        if(heading)field('b','1');
        const newId=children(fonts,'font').length;fonts.append(font);fonts.setAttribute('count',newId+1);fontMap.set(fk,newId);
      }
      set(xf,{fontId:fontMap.get(fk),applyFont:1,applyAlignment:1});
      const align=insert(xf,'alignment',['alignment','protection','extLst']);
      set(align,{wrapText:1,shrinkToFit:0,vertical:'center',textRotation:0,indent:0});
      const index=children(xfs,'xf').length;xfs.append(xf);xfs.setAttribute('count',index+1);mapped.set(key,index);return index;
    }
    return {getStyle,digitWidth};
  }
  function setColumns(root, widths, digitWidth){
    const container=insert(root,'cols',sheetOrder), prior=children(container,'col'), boundaries=new Set();
    for(const el of prior){boundaries.add(+el.getAttribute('min'));boundaries.add(+el.getAttribute('max')+1);}
    for(const c of widths.keys()){boundaries.add(c+1);boundaries.add(c+2);}
    const cuts=[...boundaries].sort((a,b)=>a-b);container.replaceChildren();
    let j=0;
    for(let i=0;i+1<cuts.length;i++){
      const a=cuts[i],b=cuts[i+1]-1;
      while(j<prior.length && +prior[j].getAttribute('max')<a)j++;
      const old=prior[j] && +prior[j].getAttribute('min')<=a?prior[j]:null;
      if(!old && !widths.has(a-1))continue;
      const col=old?old.cloneNode(true):make(root,'col');set(col,{min:a,max:b});
      if(widths.has(a-1)){set(col,{width:Math.ceil(widths.get(a-1)*96/72/digitWidth*256)/256,customWidth:1});col.removeAttribute('bestFit');}
      container.append(col);
    }
  }
  function setBreaks(root,nodeName,values,max){
    first(root,nodeName)?.remove();
    if(!values.length)return;
    const part=insert(root,nodeName,sheetOrder);set(part,{count:values.length,manualBreakCount:values.length});
    for(const id of values){const e=make(root,'brk');set(e,{id,max,man:1});part.append(e);}
  }
  async function exportBook(source, views){
    const input=source.original || XLSX.write(source.book,{type:'array',bookType:'xlsx'});
    const zip=await JSZip.loadAsync(input);
    const packageRels=await doc(zip,'_rels/.rels');
    const lookup=(d,suffix)=>Array.from(d.documentElement.children).find(el=>el.getAttribute('Type')?.endsWith('/'+suffix));
    const bookPart=partPath('',lookup(packageRels,'officeDocument').getAttribute('Target'));
    const book=await doc(zip,bookPart),rels=await doc(zip,relsPath(bookPart)),br=book.documentElement;
    const sheetList=children(first(br,'sheets'),'sheet');
    let stylesRel=lookup(rels,'styles'),stylePart,styles;
    if(stylesRel){
      stylePart=partPath(bookPart,stylesRel.getAttribute('Target'));styles=await doc(zip,stylePart);
    }else{
      const seed=fromRows([['']], 'Sheet1');
      const seedZip=await JSZip.loadAsync(XLSX.write(seed,{type:'array',bookType:'xlsx'}));
      styles=await doc(seedZip,'xl/styles.xml');
      stylePart=partPath(bookPart,'styles.xml');
      if(zip.file(stylePart))throw Error('A worksheet style part already uses the required name.');
      const relroot=rels.documentElement,rel=make(relroot,'Relationship');
      let id='fixhubPrintStyles';while(Array.from(relroot.children).some(e=>e.getAttribute('Id')===id))id+='1';
      set(rel,{Id:id,Type:relNS+'/styles',Target:'styles.xml'});relroot.append(rel);
      const types=await doc(zip,'[Content_Types].xml'),override=make(types.documentElement,'Override');
      set(override,{PartName:'/'+stylePart,ContentType:'application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml'});types.documentElement.append(override);
      save(zip,'[Content_Types].xml',types);save(zip,relsPath(bookPart),rels);
    }
    const {getStyle,digitWidth}=updateStyles(styles);
    for(const {sheet,layout} of views){
      const repeat=layout.repeat;
      const idx=sheetList.findIndex(el=>el.getAttribute('name')===sheet.name);
      const rid=sheetList[idx].getAttributeNS(relNS,'id')||sheetList[idx].getAttribute('r:id');
      const rel=Array.from(rels.documentElement.children).find(el=>el.getAttribute('Id')===rid);
      const part=partPath(bookPart,rel.getAttribute('Target'));
      const ws=await doc(zip,part),root=ws.documentElement;
      setColumns(root,layout.widths,digitWidth);
      const data=first(root,'sheetData'), rowNodes=new Map(children(data,'row').map(r=>[Number(r.getAttribute('r'))-1,r]));
      // Insert blank row records only where a height was requested, keeping XML row order.
      let following=null;
      const ordered=[...new Set([...rowNodes.keys(),...layout.heights.keys()])].sort((a,b)=>b-a);
      for(const r of ordered){
        let row=rowNodes.get(r);
        if(!row){row=make(root,'row');row.setAttribute('r',r+1);data.insertBefore(row,following);}
        following=row;
        if(!layout.heights.has(r))continue;
        set(row,{ht:Math.min(409.5,Math.ceil(layout.heights.get(r))),customHeight:1});
        for(const cell of children(row,'c')){
          const address=cell.getAttribute('r');if(!address || !layout.widths.has(XLSX.utils.decode_cell(address).c))continue;
          cell.setAttribute('s',getStyle(Number(cell.getAttribute('s'))||0,repeat && r===sheet.rows[0]));
          if(!source.original && source.book.Sheets[sheet.name][address]?.t==='s'){
            // Write imported text literally, including quoted CSV newlines and leading zeros.
            const value=source.book.Sheets[sheet.name][address].v;
            const inline=make(root,'is'),text=make(root,'t');
            text.setAttributeNS('http://www.w3.org/XML/1998/namespace','xml:space','preserve');
            text.textContent=value.replace(/_x([0-9a-f]{4})_/gi,'_x005F_x$1_')
              .replace(/[\x00-\x08\x0B\x0C\x0E-\x1F]/g,c=>'_x'+c.charCodeAt(0).toString(16).padStart(4,'0')+'_');
            inline.append(text);cell.setAttribute('t','inlineStr');cell.replaceChildren(inline);
          }
        }
      }
      const sheetPr=insert(root,'sheetPr',sheetOrder);
      set(insert(sheetPr,'pageSetUpPr',['tabColor','outlinePr','pageSetUpPr']),{fitToPage:0});
      const margins=insert(root,'pageMargins',sheetOrder);
      set(margins,{left:layout.margin/72,right:layout.margin/72,top:layout.gridTop/72,bottom:layout.margin/72,header:0.16,footer:0.16});
      const setup=insert(root,'pageSetup',sheetOrder);
      set(setup,{paperSize:layout.paperId,orientation:layout.orientation,scale:100,pageOrder:'overThenDown',usePrinterDefaults:0});
      setup.removeAttribute('fitToWidth');setup.removeAttribute('fitToHeight');
      setBreaks(root,'colBreaks',layout.colBreaks,1048575);
      setBreaks(root,'rowBreaks',layout.rowBreaks,16383);
      const defs=insert(br,'definedNames',bookOrder);
      for(const d of children(defs,'definedName'))if(d.getAttribute('localSheetId')===String(idx) && ['_xlnm.Print_Area','_xlnm.Print_Titles'].includes(d.getAttribute('name')))d.remove();
      const prefix="'"+sheet.name.replace(/'/g,"''")+"'!";
      function def(name,v){const el=make(br,'definedName');set(el,{name,localSheetId:idx});el.textContent=prefix+v;defs.append(el);}
      const col=XLSX.utils.encode_col;
      def('_xlnm.Print_Area','$'+col(sheet.cols[0])+'$'+(sheet.rows[0]+1)+':$'+col(sheet.cols.at(-1))+'$'+(sheet.rows.at(-1)+1));
      const titles=[];
      if(repeat)titles.push('$'+(sheet.rows[0]+1)+':$'+(sheet.rows[0]+1));
      if(layout.repeatKey)titles.push('$'+col(sheet.cols[0])+':$'+col(sheet.cols[0]));
      if(titles.length)def('_xlnm.Print_Titles',titles.join(','+prefix));
      save(zip,part,ws);
    }
    save(zip,stylePart,styles);save(zip,bookPart,book);
    return zip.generateAsync({type:'uint8array',compression:'DEFLATE'});
  }
  return {read,fromRows,model,exportBook};
})();
