(() => {
  'use strict';
  const $=id=>document.getElementById(id), io=window.FixHubWorkbook;
  const papers={A4:[210,297,9], A3:[297,420,8], Letter:[215.9,279.4,1], Legal:[215.9,355.6,5]};
  const pt=mm=>mm*72/25.4, margin=29, gridTop=55, fontSize=10, leading=13, pad=5;
  const ctx=document.createElement('canvas').getContext('2d');
  const svgNS='http://www.w3.org/2000/svg';
  const sizeCache=new Map(), graphemes=window.Intl?.Segmenter?new Intl.Segmenter(undefined,{granularity:'grapheme'}):null;
  let source, sheet, settings, layout, current=0, loadToken=0, states=new Map(), drag=null, zoomed=false;
  function measure(text,bold){
    const key=(bold?'b':'n')+text;
    if(sizeCache.has(key))return sizeCache.get(key);
    ctx.font=(bold?'bold ':'')+fontSize+'px Arial, sans-serif';
    const size=ctx.measureText(text).width;
    if(sizeCache.size>24000)sizeCache.clear();
    if(text.length<256)sizeCache.set(key,size);
    return size;
  }
  function wrap(text,width,bold){
    const result=[];
    for(const para of text.replace(/\r\n?/g,'\n').replace(/\t/g,'    ').split('\n')){
      let line='';
      for(const token of para.match(/\s+|\S+/gu)||[]){
        if(measure(line+token,bold)<=width){line+=token;continue;}
        if(line){result.push(line.trimEnd());line='';}
        if(!token.trim())continue;
        if(measure(token,bold)<=width){line=token;continue;}
        const chars=graphemes?Array.from(graphemes.segment(token),s=>s.segment):Array.from(token);
        for(const ch of chars){
          if(line && measure(line+ch,bold)>width){result.push(line);line='';}
          line+=ch;
        }
      }
      result.push(line);
    }
    return result;
  }
  function defaults(){return {paper:'A4',orientation:'auto',repeat:true,widths:new Map(),heights:new Map()};}
  function makeLayout(s,opt){
    const [short,long,paperId]=papers[opt.paper], natural=new Map(s.cols.map(c=>[c,43]));
    for(const [r,cells] of s.cells){
      const bold=opt.repeat && r===s.rows[0];
      for(const [c,val] of cells){
        const old=natural.get(c);if(old>=208)continue;
        let width=old;
        for(const line of val.split(/[\r\n]/))width=Math.max(width,Math.ceil(measure(line.slice(0,160).replace(/\t/g,'    '),bold))+2*pad+3);
        natural.set(c,Math.min(208,width));
      }
    }
    const total=s.cols.reduce((sum,c)=>sum+(opt.widths.get(c)??natural.get(c)),0);
    const orientation=opt.orientation==='auto'?(total<=pt(short)-2*margin?'portrait':'landscape'):opt.orientation;
    const [wmm,hmm]=orientation==='portrait'?[short,long]:[long,short],w=pt(wmm),h=pt(hmm),inner=w-2*margin;
    const widths=new Map(),bands=[];let group=[],used=0;
    // Use a little wrapping before resorting to a second horizontal page.
    const fixed=s.cols.reduce((n,c)=>n+(opt.widths.has(c)?Math.min(inner,opt.widths.get(c)):0),0);
    const flexible=s.cols.reduce((n,c)=>n+(opt.widths.has(c)?0:natural.get(c)),0);
    const shrink=total>inner && total<=inner*1.22 && flexible?Math.min(1,Math.max(0,(inner-fixed)/flexible)):1;
    for(const c of s.cols){
      const width=Math.min(inner,Math.max(35,opt.widths.get(c)??Math.floor(natural.get(c)*shrink)));
      widths.set(c,width);
    }
    const key=s.cols[0],keyWidth=widths.get(key)||0;
    const canRepeatKey=s.cols.slice(1).every(c=>keyWidth+widths.get(c)<=inner+.01);
    const colBreaks=[];
    for(const c of s.cols){
      const width=widths.get(c);
      if(group.length && used+width>inner+.01){
        bands.push(group);colBreaks.push(c);
        group=canRepeatKey?[key]:[];used=canRepeatKey?keyWidth:0;
      }
      group.push(c);used+=width;
    }
    if(group.length)bands.push(group);
    const lines=new Map(),heights=new Map(),minimums=new Map(),counts=new Map();
    for(const r of s.rows){
      let count=1;
      const data=s.cells.get(r),lineMap=new Map();
      if(data)for(const [c,value] of data){
        const parts=wrap(value,widths.get(c)-2*pad-1,opt.repeat && r===s.rows[0]);
        count=Math.max(count,parts.length);lineMap.set(c,parts);
      }
      if(lineMap.size)lines.set(r,lineMap);
      counts.set(r,count);
      const min=count*leading+2*pad;minimums.set(r,min);
      heights.set(r,Math.max(min,Math.min(h-gridTop-margin,opt.heights.get(r)||0)));
    }
    const bodyHeight=h-gridTop-margin, header=s.rows[0];
    const repeat=Boolean(opt.repeat && header!=null && heights.get(header)<=bodyHeight*.4);
    const headerH=repeat?heights.get(header):0, capacity=bodyHeight-headerH;
    const vertical=[];let segments=[],free=capacity;
    function finish(){vertical.push(segments);segments=[];free=capacity;}
    for(const r of s.rows){
      if(repeat&&r===header)continue;
      const height=heights.get(r),count=counts.get(r);
      if(height<=capacity+.001){
        if(free+.001<height && segments.length)finish();
        segments.push({r,from:0,to:count,height});free-=height;continue;
      }
      let from=0;
      while(from<count){
        const fit=Math.floor((free-2*pad+.001)/leading);
        if(fit<1){finish();continue;}
        const to=Math.min(count,from+fit),partHeight=(to-from)*leading+2*pad;
        segments.push({r,from,to,height:partHeight});free-=partHeight;from=to;
        if(from<count)finish();
      }
    }
    if(segments.length||(!vertical.length&&s.rows.length))finish();
    const rowBreaks=vertical.slice(1).filter(page=>page[0]?.from===0).map(page=>page[0].r);
    return {w,h,wmm,hmm,paperId,orientation,margin,gridTop,bands,vertical,repeat,header,headerH,counts,
      widths,heights,minimums,lines,rowBreaks,colBreaks,repeatKey:canRepeatKey&&bands.length>1,pageCount:vertical.length*bands.length};
  }
  function svg(type,attrs={},value){
    const node=document.createElementNS(svgNS,type);
    for(const [k,v] of Object.entries(attrs))node.setAttribute(k,v);
    if(value!==undefined)node.textContent=value;
    return node;
  }
  function shortLabel(value,max,bold){
    if(measure(value,bold)<=max)return value;
    const chars=Array.from(value);let low=0,high=chars.length;
    while(low<high){const mid=Math.ceil((low+high)/2);if(measure(chars.slice(0,mid).join('')+'…',bold)<=max)low=mid;else high=mid-1;}
    return chars.slice(0,low).join('')+'…';
  }
  function pageSVG(page,interactive){
    const l=layout,band=l.bands[page%l.bands.length],parts=l.vertical[Math.floor(page/l.bands.length)];
    const root=svg('svg',{xmlns:svgNS,viewBox:'0 0 '+l.w+' '+l.h,width:l.w,height:l.h,role:'group','aria-label':'Print preview, page '+(page+1),'font-family':'Arial, sans-serif','font-size':fontSize,fill:'#242337'});
    root.append(svg('rect',{width:l.w,height:l.h,fill:'white'}));
    root.append(svg('text',{x:margin,y:30,'font-weight':'bold','font-size':13},shortLabel(sheet.name,l.w*.54/1.4,true)));
    root.append(svg('text',{x:l.w-margin,y:30,'text-anchor':'end',fill:'#625c72','font-size':8.5},shortLabel(source.name,l.w*.34,false)));
    root.append(svg('line',{x1:margin,x2:l.w-margin,y1:38,y2:38,stroke:'#825fc5','stroke-width':1.3}));
    const paperRows=l.repeat?[{r:l.header,from:0,to:l.counts.get(l.header),height:l.headerH},...parts]:parts;
    let y=gridTop;
    const controls=[];
    function grip(kind,index,x,y,width,height){
      const isCol=kind==='col',prefix=isCol?'Column '+XLSX.utils.encode_col(index):'Row '+(index+1);
      const v=(isCol?l.widths:l.heights).get(index);
      const rect=svg('rect',{x,y,width,height,rx:2,class:'resize '+kind+'-resize','data-axis':kind,'data-index':index,tabindex:0,role:'separator','aria-orientation':isCol?'vertical':'horizontal','aria-label':prefix+' size, drag or use arrow keys','aria-valuenow':Math.round(v),'aria-valuetext':Math.round(v)+' points'});
      rect.append(svg('title',{},prefix+' · drag or use arrow keys'));
      controls.push(rect);
    }
    if(interactive){
      let x=margin;
      for(const c of band){
        x+=l.widths.get(c);
        root.append(svg('path',{d:'M '+(x-2)+' '+(gridTop-7)+' h 4 m -2 -3 v 8',stroke:'#aa9bbf','stroke-width':.8,fill:'none'}));
        grip('col',c,x-5,gridTop-14,10,14+Math.min(24,paperRows[0].height));
      }
    }
    for(const p of paperRows){
      const bold=settings.repeat&&p.r===sheet.rows[0], lineMap=l.lines.get(p.r);
      let x=margin;
      for(const c of band){
        const cw=l.widths.get(c);
        root.append(svg('rect',{x,y,width:cw,height:p.height,fill:bold?'#f0edf7':'white',stroke:'#c9c7d0','stroke-width':.55}));
        const rows=lineMap?.get(c);
        if(rows){
          const text=svg('text',{'font-weight':bold?'bold':'normal','xml:space':'preserve',style:'white-space:pre'});
          for(let i=p.from;i<Math.min(p.to,rows.length);i++){
            text.append(svg('tspan',{x:x+pad,y:y+pad+fontSize+(i-p.from)*leading},rows[i]));
          }
          root.append(text);
        }
        x+=cw;
      }
      y+=p.height;
      if(interactive){
        root.append(svg('path',{d:'M '+(margin-9)+' '+y+' h 7',stroke:'#aa9bbf','stroke-width':.9}));
        grip('row',p.r,margin-20,y-5,30,10);
      }
    }
    root.append(svg('text',{x:l.w-margin,y:l.h-13,'text-anchor':'end',fill:'#686372','font-size':8},'Page '+(page+1)+' of '+l.pageCount));
    if(interactive)root.append(...controls);
    return root;
  }
  function status(message,error=false){$('message').textContent=message;$('message').classList.toggle('error',error);}
  function show(){
    $('print-stack').replaceChildren();
    current=Math.max(0,Math.min(current,layout.pageCount-1));
    const preview=$('preview');preview.replaceChildren();
    if(layout.pageCount){
      const node=pageSVG(current,true);node.classList.add('sheet-preview');
      if(zoomed){node.style.width=Math.round(layout.w*1.65)+'px';node.style.maxWidth='none';}
      else node.style.maxWidth=Math.min(1000,layout.w*1.28)+'px';
      preview.append(node);
    }else{const p=document.createElement('p');p.textContent='This worksheet has no visible cell data.';preview.append(p);}
    $('filename').textContent=source.name;
    $('counter').textContent=layout.pageCount?'Page '+(current+1)+' of '+layout.pageCount:'0 pages';
    $('prev').disabled=current===0;$('next').disabled=current+1>=layout.pageCount;
    $('export').disabled=$('print').disabled=!layout.pageCount;
    $('paper-style').textContent='@page{size:'+layout.wmm+'mm '+layout.hmm+'mm;margin:0} @media print{#print-stack .printed{width:'+layout.wmm+'mm;height:'+(layout.hmm-.05)+'mm}}';
    const details=[settings.paper,layout.orientation==='portrait'?'Portrait':'Landscape'];
    if(layout.bands.length>1)details.push(layout.bands.length+' pages across'+(layout.repeatKey?' with the first column repeated':''));
    if(sheet.uncached)details.push('Uncalculated formulas are shown as formulas');
    if(settings.repeat && layout.pageCount && !layout.repeat)details.push('First row too tall to repeat');
    if([...layout.heights.values()].some(h=>h>409.5))details.push('For very tall rows, use Print / Save PDF to include every line');
    $('size-info').textContent=details.join(' · ');
  }
  function refresh(){layout=makeLayout(sheet,settings);show();}
  function selectSheet(){
    const name=$('sheet').value;sheet=io.model(source.book,name);
    if(!states.has(name))states.set(name,defaults());
    settings=states.get(name);current=0;
    $('paper').value=settings.paper;$('orientation').value=settings.orientation;$('repeat').checked=settings.repeat;
    refresh();
  }
  function install(next){
    const visible=next.book.SheetNames.filter((_,i)=>!next.book.Workbook?.Sheets?.[i]?.Hidden);
    if(!visible.length)throw Error('There are no visible worksheets in this workbook.');
    const first=visible.find(name=>next.book.Sheets[name]['!ref'])||visible[0];
    const prior={source,states,sheet,settings,layout,current,names:Array.from($('sheet').options,o=>o.value),selected:$('sheet').value};
    try{
      source=next;states=new Map();$('sheet').replaceChildren(...visible.map(n=>new Option(n,n)));$('sheet').value=first;selectSheet();
    }catch(e){
      ({source,states,sheet,settings,layout,current}=prior);
      $('sheet').replaceChildren(...prior.names.map(n=>new Option(n,n)));$('sheet').value=prior.selected;
      if(layout)show();throw e;
    }
  }
  async function open(file){
    if(!file)return;
    const token=++loadToken;status('Preparing '+file.name+'…');
    try{const next=await io.read(file);if(token!==loadToken)return;install(next);status('');}
    catch(e){if(token===loadToken)status('Could not open '+file.name+'. '+e.message,true);}
  }
  $('file').onchange=e=>{open(e.target.files[0]);e.target.value='';};
  $('editor').ondragover=e=>{if(e.dataTransfer.types.includes('Files')){e.preventDefault();$('editor').classList.add('dragging');}};
  $('editor').ondragleave=()=>$('editor').classList.remove('dragging');
  $('editor').ondrop=e=>{e.preventDefault();$('editor').classList.remove('dragging');open(e.dataTransfer.files[0]);};
  $('sheet').onchange=()=>{status('');selectSheet();};
  for(const id of ['paper','orientation','repeat'])$(id).onchange=()=>{settings[id]=id==='repeat'?$(id).checked:$(id).value;current=0;refresh();};
  $('zoom').onclick=()=>{
    zoomed=!zoomed;$('preview').classList.toggle('zoomed',zoomed);
    $('zoom').setAttribute('aria-pressed',String(zoomed));$('zoom').textContent=zoomed?'Fit page':'Zoom in';
    $('zoom').setAttribute('aria-label',zoomed?'Fit the whole page to the preview':'Zoom in on the print preview');show();
  };
  $('prev').onclick=()=>{current--;show();};$('next').onclick=()=>{current++;show();};
  $('reset').onclick=()=>{states.delete(sheet.name);selectSheet();status('Automatic layout restored for this worksheet.');};
  function resize(axis,index,value){
    if(axis==='col')settings.widths.set(index,Math.round(Math.min(layout.w-2*margin,Math.max(35,value))));
    else settings.heights.set(index,Math.max(layout.minimums.get(index),Math.min(layout.h-gridTop-margin,value)));
    refresh();
  }
  const preview=$('preview');
  preview.addEventListener('pointerdown',e=>{
    const handle=e.target.closest('[data-axis]');if(!handle || (e.button!==0 && e.pointerType==='mouse'))return;
    const axis=handle.dataset.axis,index=Number(handle.dataset.index);
    drag={axis,index,x:e.clientX,y:e.clientY,id:e.pointerId,scale:layout.w/handle.ownerSVGElement.getBoundingClientRect().width,
      initial:(axis==='col'?layout.widths:layout.heights).get(index)};
    e.preventDefault();preview.setPointerCapture(e.pointerId);
  });
  let pending=0,latest=null;
  function applyDrag(e){if(!drag)return;const delta=drag.axis==='col'?e.clientX-drag.x:e.clientY-drag.y;resize(drag.axis,drag.index,drag.initial+delta*drag.scale);}
  preview.addEventListener('pointermove',e=>{
    if(!drag||e.pointerId!==drag.id)return;latest=e;
    if(!pending)pending=requestAnimationFrame(()=>{pending=0;if(latest)applyDrag(latest);});
  });
  function endDrag(e){
    if(!drag||e.pointerId!==drag.id)return;
    cancelAnimationFrame(pending);pending=0;latest=null;
    if(e.type==='pointerup')applyDrag(e);
    drag=null;if(preview.hasPointerCapture(e.pointerId))preview.releasePointerCapture(e.pointerId);
  }
  preview.addEventListener('pointerup',endDrag);preview.addEventListener('pointercancel',endDrag);
  preview.addEventListener('keydown',e=>{
    const el=e.target.closest('[data-axis]');if(!el)return;
    const axis=el.dataset.axis, keys=axis==='col'?['ArrowLeft','ArrowRight']:['ArrowUp','ArrowDown'];
    if(!keys.includes(e.key))return;e.preventDefault();
    const index=Number(el.dataset.index),value=(axis==='col'?layout.widths:layout.heights).get(index);
    resize(axis,index,value+(e.key===keys[1]?1:-1)*(e.shiftKey?12:4));
    preview.querySelector('[data-axis="'+axis+'"][data-index="'+index+'"]')?.focus();
  });
  function preparePrint(){
    const fragment=document.createDocumentFragment();
    for(let p=0;p<layout.pageCount;p++){const div=document.createElement('div');div.className='printed';div.append(pageSVG(p,false));fragment.append(div);}
    $('print-stack').replaceChildren(fragment);
  }
  $('print').onclick=()=>{preparePrint();window.print();};
  window.addEventListener('beforeprint',()=>{if(!$('print-stack').childElementCount)preparePrint();});
  window.addEventListener('afterprint',()=>$('print-stack').replaceChildren());
  $('export').onclick=async()=>{
    const btn=$('export');btn.disabled=true;
    const name=source.name.replace(/\.(csv|xlsx)$/i,'')+'-print.xlsx';
    try{
      const views=[];
      for(const [name,opt] of states){
        if(name===sheet.name){if(layout.pageCount)views.push({sheet,layout});continue;}
        const model=io.model(source.book,name);
        if(model.rows.length)views.push({sheet:model,layout:makeLayout(model,opt)});
      }
      const bytes=await io.exportBook(source,views);
      const blob=new Blob([bytes],{type:'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'});
      const url=URL.createObjectURL(blob),link=document.createElement('a');link.href=url;link.download=name;document.body.append(link);link.click();link.remove();
      setTimeout(()=>URL.revokeObjectURL(url),30000);status('Downloaded '+name+'.');
    }catch(e){status('Could not create the Excel download. '+e.message,true);}
    finally{btn.disabled=!layout.pageCount;}
  };
  document.addEventListener('click',event=>{
    if(event.detail!==1 || event.button!==0 || event.shiftKey || event.target.closest('.resize') || event.target.matches(':read-write'))return;
    const selection=window.getSelection();
    if(selection?.rangeCount && selection.isCollapsed)selection.removeAllRanges();
  },true);
  const sample=[
    ['Order','Customer','Item description','Quantity','Delivery date'],
    ['SO-1042','Willow Studio','Natural oak desk with two drawers',2,'24 November'],
    ['SO-1043','North House','Adjustable reading lamp, brushed metal finish',4,'25 November'],
    ['SO-1044','Field & Form','Oak shelving with wall mounting brackets',1,'26 November'],
    ['SO-1045','River Lane','Linen-covered noticeboard',3,'27 November'],
    ['SO-1046','Ash Works','Stacking chair, charcoal finish',6,'30 November']
  ];
  try{install({name:'Sample worksheet',book:io.fromRows(sample,'Orders'),original:null});}
  catch(e){status('The formatter could not start. '+e.message,true);}
})();
