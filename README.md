<!DOCTYPE html>
<html lang="ar" dir="rtl">
<head>
<meta charset="utf-8">
<meta name="viewport" content="width=device-width, initial-scale=1, viewport-fit=cover">
<title>مُستخرج أوراق إكسل</title>
<style>
:root{--bg:#f6f7fb;--card:#fff;--text:#1b1f2a;--muted:#667085;--line:#e3e6ee;--accent:#1f7a4d;--accent-t:#fff;--soft:#e7f4ec;--err:#b42318;
box-sizing:border-box;padding-top:env(safe-area-inset-top,0px);padding-bottom:env(safe-area-inset-bottom,0px)}
@media(prefers-color-scheme:dark){:root:not([data-theme="light"]){--bg:#12151c;--card:#1b2029;--text:#eef0f5;--muted:#9aa3b5;--line:#2c3340;--accent:#3fb97c;--accent-t:#08150e;--soft:#16281f;--err:#f97066}}
:root[data-theme="dark"]{--bg:#12151c;--card:#1b2029;--text:#eef0f5;--muted:#9aa3b5;--line:#2c3340;--accent:#3fb97c;--accent-t:#08150e;--soft:#16281f;--err:#f97066}
*{box-sizing:border-box}
body{margin:0;background:var(--bg);color:var(--text);font-family:"Segoe UI",Tahoma,"Noto Sans Arabic",system-ui,sans-serif;line-height:1.6}
main{max-width:640px;margin:0 auto;padding:24px 16px 48px}
h1{font-size:1.5rem;margin:0 0 4px}
.sub{color:var(--muted);margin:0 0 20px}
.drop{display:block;border:2px dashed var(--line);background:var(--card);border-radius:16px;padding:32px 16px;text-align:center;cursor:pointer;transition:.15s}
.drop:hover,.drop.over{border-color:var(--accent);background:var(--soft)}
.drop b{display:block;font-size:1.05rem}
.drop span{color:var(--muted);font-size:.9rem}
input[type=file]{display:none}
.card{background:var(--card);border:1px solid var(--line);border-radius:16px;padding:16px;margin-top:16px}
.row{display:flex;align-items:center;gap:10px;flex-wrap:wrap}
.grow{flex:1;min-width:0}
.seg{display:inline-flex;border:1px solid var(--line);border-radius:10px;overflow:hidden}
.seg button{border:0;background:transparent;color:var(--text);padding:6px 14px;cursor:pointer;font:inherit}
.seg button.on{background:var(--accent);color:var(--accent-t)}
.btn{border:0;background:var(--accent);color:var(--accent-t);padding:8px 16px;border-radius:10px;font:inherit;cursor:pointer}
.btn.ghost{background:var(--soft);color:var(--accent)}
ul{list-style:none;margin:6px 0 0;padding:0}
li{display:flex;align-items:center;gap:10px;padding:8px 0;border-top:1px solid var(--line)}
li:first-child{border-top:0}
.name{font-weight:600;overflow:hidden;text-overflow:ellipsis;white-space:nowrap}
.meta{color:var(--muted);font-size:.85rem}
.book{margin-top:16px;border:1px solid var(--line);border-radius:12px;padding:10px 12px}
#msg{margin-top:12px;min-height:1.4em;font-size:.92rem;color:var(--muted)}
#msg.err{color:var(--err)}
.hidden{display:none!important}
</style>
</head>
<body>
<main>
  <h1>مُستخرج أوراق إكسل</h1>
  <p class="sub">ارفع ملف إكسل واحدًا أو أكثر، ثم احفظ كل ورقة في ملف مستقل أو ادمجها كلها في CSV واحد. المعالجة تتم داخل متصفحك فقط.</p>

  <label class="drop" id="drop">
    <b>اختر ملفًا أو أكثر، أو اسحبها هنا</b>
    <span>.xlsx · .xlsm · .xls · .csv — يمكنك إضافة ملفات في أي وقت</span>
    <input type="file" id="file" multiple accept=".xlsx,.xlsm,.xls,.xlsb,.csv">
  </label>

  <section class="card hidden" id="result">
    <div class="row">
      <div class="grow name" id="summary"></div>
      <div class="seg" id="fmt"><button data-f="csv" class="on">CSV</button><button data-f="xlsx">XLSX</button></div>
    </div>
    <div id="books"></div>
    <div class="row" style="margin-top:14px">
      <button class="btn grow" id="mergeBtn">دمج كل الأوراق في ملف CSV واحد</button>
      <button class="btn ghost" id="zipBtn">ZIP منفصل</button>
    </div>
    <label class="meta" style="display:block;margin-top:10px"><input type="checkbox" id="addCol" checked> إضافة أعمدة المصدر (الملف والورقة) لكل صف</label>
    <button class="btn ghost" id="clearBtn" style="margin-top:10px">مسح كل الملفات</button>
  </section>
  <div id="msg" role="status"></div>
</main>

<script src="https://cdnjs.cloudflare.com/ajax/libs/xlsx/0.18.5/xlsx.full.min.js"></script>
<script src="https://cdnjs.cloudflare.com/ajax/libs/jszip/3.10.1/jszip.min.js"></script>
<script>
(function(){
  var books=[], fmt="csv", downloads=null;
  var $=function(id){return document.getElementById(id)};
  var msg=$("msg");

  function say(t,err){msg.textContent=t||"";msg.className=err?"err":""}
  function safeName(n){var s=String(n).replace(/[\\\/:*?"<>|]+/g,"_").trim();return s||"sheet"}
  function stem(n){return safeName(String(n).replace(/\.[^.]+$/,""))}

  // One entry per sheet across all workbooks, with unique output file names
  function plan(){
    var used={},items=[],multi=books.length>1;
    books.forEach(function(b,bi){
      b.wb.SheetNames.forEach(function(n){
        var base=(multi?stem(b.name)+"_":"")+safeName(n),c=base,i=1;
        while(used[c.toLowerCase()]){i++;c=base+"_"+i}
        used[c.toLowerCase()]=1;
        items.push({bi:bi,sheet:n,file:c+"_data."+fmt});
      });
    });
    return items;
  }

  function sheetData(it){
    var ws=books[it.bi].wb.Sheets[it.sheet];
    if(fmt==="csv") return "\uFEFF"+XLSX.utils.sheet_to_csv(ws);
    var nb=XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(nb,ws,String(it.sheet).slice(0,31));
    return XLSX.write(nb,{bookType:"xlsx",type:"array"});
  }

  function rowCount(it){
    var ref=books[it.bi].wb.Sheets[it.sheet]["!ref"];
    if(!ref) return 0;
    var r=XLSX.utils.decode_range(ref);
    return r.e.r-r.s.r+1;
  }

  async function save(filename,data){
    if(!downloads){say("الحفظ غير متاح في هذا العرض.",true);return}
    try{
      await downloads.save({filename:filename,data:data});
      say("تم الحفظ: "+filename);
    }catch(e){
      if(e&&e.code==="declined") say("تم إلغاء الحفظ.");
      else say("تعذّر الحفظ ("+((e&&e.code)||"خطأ")+").",true);
    }
  }

  function render(){
    var box=$("books");
    box.innerHTML="";
    if(!books.length){$("result").classList.add("hidden");return}
    var items=plan();
    $("summary").textContent=books.length+" ملف · "+items.length+" ورقة";
    books.forEach(function(b,bi){
      var g=document.createElement("div");g.className="book";
      var h=document.createElement("div");h.className="row";
      var t=document.createElement("div");t.className="grow name";t.textContent=b.name;
      var x=document.createElement("button");x.className="btn ghost";x.textContent="إزالة";
      x.onclick=function(){books.splice(bi,1);render();say("")};
      h.appendChild(t);h.appendChild(x);g.appendChild(h);
      var ul=document.createElement("ul");
      items.filter(function(it){return it.bi===bi}).forEach(function(it){
        var li=document.createElement("li");
        var info=document.createElement("div");info.className="grow";
        var a=document.createElement("div");a.className="name";a.textContent=it.sheet;
        var m=document.createElement("div");m.className="meta";m.textContent=rowCount(it)+" صف · "+it.file;
        info.appendChild(a);info.appendChild(m);
        var btn=document.createElement("button");btn.className="btn ghost";btn.textContent="تحميل";
        btn.onclick=function(){save(it.file,sheetData(it))};
        li.appendChild(info);li.appendChild(btn);ul.appendChild(li);
      });
      g.appendChild(ul);box.appendChild(g);
    });
    $("result").classList.remove("hidden");
  }

  async function zipAll(){
    var zip=new JSZip();
    plan().forEach(function(it){zip.file(it.file,sheetData(it))});
    var blob=await zip.generateAsync({type:"blob"});
    save((books.length>1?"workbooks":stem(books[0].name))+"_sheets.zip",blob);
  }

  function csvCell(v){
    v=String(v);
    return /[",\r\n]/.test(v)?'"'+v.replace(/"/g,'""')+'"':v;
  }

  // Stack every sheet of every workbook into one table; columns are matched by header name
  function mergeAll(){
    var addCol=$("addCol").checked,multi=books.length>1;
    var cols=[],idx=Object.create(null),recs=[];
    books.forEach(function(b){
      b.wb.SheetNames.forEach(function(n){
        var rows=XLSX.utils.sheet_to_json(b.wb.Sheets[n],{header:1,raw:false,defval:""});
        if(!rows.length) return;
        var seen=Object.create(null);
        var head=rows[0].map(function(h,i){
          var s=String(h).trim()||("عمود "+(i+1)),c=s,k=1;
          while(seen[c]){k++;c=s+" ("+k+")"}
          seen[c]=1;return c;
        });
        head.forEach(function(h){if(!(h in idx)){idx[h]=1;cols.push(h)}});
        for(var r=1;r<rows.length;r++){
          var o=Object.create(null),has=false;
          head.forEach(function(h,i){
            var v=rows[r][i]===undefined?"":rows[r][i];
            if(String(v)!=="") has=true;
            o[h]=v;
          });
          if(has) recs.push({f:b.name,s:n,o:o});
        }
      });
    });
    if(!recs.length){say("لا توجد بيانات للدمج.",true);return}
    var pre=addCol?(multi?["الملف","الورقة"]:["الورقة"]):[];
    var lines=[pre.concat(cols).map(csvCell).join(",")];
    recs.forEach(function(rec){
      var p=addCol?(multi?[rec.f,rec.s]:[rec.s]):[];
      lines.push(p.concat(cols.map(function(c){return rec.o[c]===undefined?"":rec.o[c]})).map(csvCell).join(","));
    });
    say("تم دمج "+recs.length+" صف.");
    save(multi?"merged_workbooks.csv":stem(books[0].name)+"_merged.csv","\uFEFF"+lines.join("\r\n"));
  }

  function readFile(f){
    return new Promise(function(res){
      var rd=new FileReader();
      rd.onload=function(){
        try{res({name:f.name,wb:XLSX.read(new Uint8Array(rd.result),{type:"array"})})}
        catch(e){res({err:f.name})}
      };
      rd.onerror=function(){res({err:f.name})};
      rd.readAsArrayBuffer(f);
    });
  }

  async function loadFiles(list){
    list=Array.prototype.slice.call(list||[]);
    if(!list.length) return;
    say("جارٍ القراءة…");
    var rs=await Promise.all(list.map(readFile)),bad=[];
    rs.forEach(function(r){if(r.err)bad.push(r.err);else books.push(r)});
    render();
    say(bad.length?"تعذّرت قراءة: "+bad.join("، "):"",bad.length>0);
  }

  $("file").onchange=function(e){loadFiles(e.target.files);e.target.value=""};
  var drop=$("drop");
  ["dragenter","dragover"].forEach(function(ev){drop.addEventListener(ev,function(e){e.preventDefault();drop.classList.add("over")})});
  ["dragleave","drop"].forEach(function(ev){drop.addEventListener(ev,function(e){e.preventDefault();drop.classList.remove("over")})});
  drop.addEventListener("drop",function(e){loadFiles(e.dataTransfer.files)});

  $("fmt").onclick=function(e){
    var f=e.target.getAttribute&&e.target.getAttribute("data-f");
    if(!f||f===fmt) return;
    fmt=f;
    Array.prototype.forEach.call($("fmt").children,function(b){b.classList.toggle("on",b.getAttribute("data-f")===fmt)});
    render();
  };
  $("zipBtn").onclick=zipAll;
  $("mergeBtn").onclick=mergeAll;
  $("clearBtn").onclick=function(){books=[];render();say("")};

  if(window.claude&&claude.use){
    claude.use("downloads").then(function(d){downloads=d;if(!d)say("الحفظ المباشر غير متاح في هذا العرض.",true)}).catch(function(){});
  }
})();
</script>
</body>
</html>
