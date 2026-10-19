import fs from 'node:fs';
import path from 'node:path';
import pptxgen from 'pptxgenjs';

const root = path.resolve(import.meta.dirname, '..');
const rows = fs.readFileSync(path.join(root, 'data/regional_top7_2026_mat.csv'), 'utf8')
  .trim().split(/\r?\n/).slice(1).map(line => {
    const c = line.split(',');
    return { period:c[0], metric:c[1], region:c[2], rank:Number(c[3]), brand:c[4], source_id:c[5], limitation:c[6] };
  });
const regions = ['수도권','충청권','전라권','경상권','강원권'];
const short = new Map([['수도권','수도권'],['충청권','충청'],['전라권','전라'],['경상권','경상'],['강원권','강원']]);
const winners = regions.map(region => ({ region, brand:rows.find(x=>x.region===region&&x.rank===1).brand }));
const counts = Object.groupBy ? Object.groupBy(winners, x=>x.brand) : winners.reduce((a,x)=>((a[x.brand]??=[]).push(x),a),{});
const majority = Object.entries(counts).sort((a,b)=>b[1].length-a[1].length);
const sourceUrl = 'https://www.trendmonitor.co.kr/tmweb/trend/etcTrend/detail.do?bIdx=36&code=DDB&trendType=ranking';
const sourceFigure = 'https://trendmonitor.co.kr/Data//CKOREA/1782902151431.png';
const out = path.join(root,'presentation'); fs.mkdirSync(out,{recursive:true});
const pptx = new pptxgen(); pptx.layout='LAYOUT_WIDE'; pptx.author='Data Literacy Project'; pptx.subject='2026년 5월 MAT 지역별 가정용 소주 구매액 순위'; pptx.title='지역별 가정용 소주 구매액 1위 브랜드'; pptx.company='Data Literacy'; pptx.lang='ko-KR';
pptx.theme={headFontFace:'Aptos Display',bodyFontFace:'Malgun Gothic',lang:'ko-KR'};
const C={bg:'101A1B',panel:'1A2929',cream:'F5F0E5',ink:'152222',muted:'A6B2AA',muted2:'61706B',lime:'C9F36A',coral:'FF785E',blue:'81C8D8',line:'344341',white:'FFFFFF',gray:'DFE4DB'};
const W=13.333,H=7.5;
function box(s,x,y,w,h,fill,r=0.12,line=fill){s.addShape(pptx.ShapeType.roundRect,{x,y,w,h,rectRadius:r,line:{color:line,transparency:line===fill?100:0},fill:{color:fill},radius:r});}
function text(s,t,x,y,w,h,{size=18,color=C.ink,bold=false,fontFace='Malgun Gothic',align='left',valign='mid',margin=0,breakLine=false,transparency=0}={}){s.addText(t,{x,y,w,h,fontFace,fontSize:size,color,bold,margin,breakLine,align,valign,fit:'shrink',transparency,paraSpaceAfterPt:4,lang:'ko-KR'});}
function base(kicker,title,sub,dark=false){const s=pptx.addSlide();s.background={color:dark?C.bg:C.cream};text(s,kicker.toUpperCase(),.62,.35,10,.25,{size:10,bold:true,color:dark?C.lime:C.muted2,fontFace:'Aptos'});text(s,title,.62,.8,12.1,.62,{size:27,bold:true,color:dark?C.cream:C.ink});if(sub)text(s,sub,.64,1.5,11.9,.52,{size:12,color:dark?C.muted:C.muted2});text(s,String(pptx._slides.length).padStart(2,'0')+' / 10',12.08,7.08,.72,.2,{size:9,color:dark?C.muted:C.muted2,fontFace:'Aptos',align:'right'});return s;}
function foot(s,txt='출처: 트렌드모니터·엠브레인 DD.B, 2026년 5월 MAT | 수집·전사 확인: 2026-10-19'){text(s,txt,.64,7.04,11.2,.22,{size:8,color:C.muted2});}
function pill(s,label,x,y,w,fill,color=C.ink){box(s,x,y,w,.36,fill,.1);text(s,label,x+.08,y+.02,w-.16,.3,{size:9,bold:true,color,align:'center'});}
// 1 Cover
{
 const s=pptx.addSlide();s.background={color:C.bg};
 s.addShape(pptx.ShapeType.arc,{x:8.05,y:-.8,w:6.3,h:8.5,line:{color:C.line,width:30,transparency:30},adjustPoint:.3,rotate:22});
 box(s,.72,.66,1.26,.39,C.lime);text(s,'DATA / 2026',.82,.72,1.05,.2,{size:10,bold:true,color:C.ink,fontFace:'Aptos'});
 text(s,'지역별 가정용 소주\n구매액 1위는?',.72,1.66,8.1,1.9,{size:38,bold:true,color:C.cream});
 text(s,'5개 권역의 공개 TOP 7 순위로 확인한\n우리 동네의 1위 브랜드',.76,3.86,7.1,.72,{size:18,color:C.muted});
 text(s,'2026년 5월 MAT · 가정 내 구매액 기준',.76,5.5,6,.34,{size:12,bold:true,color:C.lime});
 // prominent bottle-inspired ranking marks
 for(let i=0;i<5;i++){box(s,9.1+i*.55,2.15,.34,2.75,[C.lime,C.coral,C.blue,C.lime,C.lime][i],.12);box(s,9.18+i*.55,1.82,.18,.43,C.cream,.08);}
 text(s,'4 / 5',9.0,5.15,3.6,.82,{size:44,bold:true,color:C.lime,fontFace:'Aptos Display'});text(s,'권역에서 같은 브랜드가 1위',9.08,6.02,3.6,.38,{size:13,color:C.cream});
 text(s,'데이터 리터러시 프로젝트  |  2026.10',.74,7.05,5,.2,{size:8,color:C.muted});
}
// 2 Answer first
{
 const s=base('THE ANSWER','1위는 한 브랜드에 모였을까?','공개된 5개 권역의 가정용 소주 구매액 순위에서 확인한 결과',true);
 box(s,.7,2.35,5.7,3.55,C.panel);text(s,'4 / 5',1.02,2.66,4.8,1.25,{size:66,bold:true,color:C.lime,fontFace:'Aptos Display'});text(s,'권역의 1위',1.08,4.02,4.5,.5,{size:24,bold:true,color:C.cream});text(s,'수도권 · 충청권 · 전라권 · 강원권',1.08,4.77,4.85,.5,{size:15,color:C.muted});pill(s,'참이슬 후레쉬',1.08,5.35,1.8,C.lime);
 box(s,6.8,2.35,5.8,3.55,C.panel);text(s,'1 / 5',7.14,2.66,4.8,1.25,{size:66,bold:true,color:C.coral,fontFace:'Aptos Display'});text(s,'다른 1위',7.2,4.02,4.5,.5,{size:24,bold:true,color:C.cream});text(s,'경상권',7.2,4.77,4.85,.5,{size:15,color:C.muted});pill(s,'진로',7.2,5.35,1.0,C.coral);
 text(s,'이 자료가 말하는 1위는 “가정용 소주류 구매액 순위”의 1위입니다.',.77,6.27,11.6,.38,{size:15,bold:true,color:C.cream});foot(s);
}
// 3 question
{
 const s=base('RESEARCH QUESTION','“우리 지역 소주”가 집에서도 1위일까?','지역 브랜드의 이미지를 실제 구매 순위와 나란히 놓고 질문을 좁혔습니다.');
 const questions=[['01','어떤 브랜드가 1위인가?','2026년 5월 MAT, 권역별 TOP 7'],['02','지역마다 1위가 달라질까?','수도권·충청·전라·경상·강원'],['03','순위가 말해 주지 않는 것은?','구매액 규모, 점유율, 외식 구매']];
 questions.forEach((q,i)=>{let x=.72+i*4.18;box(s,x,2.55,3.85,2.4,C.white);text(s,q[0],x+.24,2.82,.8,.35,{size:15,bold:true,color:C.coral,fontFace:'Aptos'});text(s,q[1],x+.24,3.42,3.35,.7,{size:20,bold:true});text(s,q[2],x+.24,4.35,3.35,.38,{size:11,color:C.muted2});});
 box(s,.72,5.45,11.9,.75,'E9E4D8');text(s,'가설: 지역별 1위는 다를 수 있지만, 전국 유통 브랜드가 여러 권역의 선두일 가능성도 있다.',.98,5.63,11.4,.36,{size:14,bold:true});foot(s);
}
// 4 Data and method
{
 const s=base('DATA / METHOD','우리가 비교한 것은 ‘가정 내 구매액 순위’','브랜드 선호도 조사나 주류 전체 소비량과는 측정 대상이 다릅니다.',true);
 const cards=[['자료 제공','트렌드모니터\n엠브레인 구매딥데이터 DD.B'],['관측 기간','2026년 5월 MAT\n최근 12개월'],['대상 범위','가정용 소주류\n전국 + 5개 권역'],['공개 항목','권역별 TOP 7\n순위와 브랜드명']];
 cards.forEach((a,i)=>{let x=.72+i*3.13;box(s,x,2.5,2.82,1.72,C.panel);text(s,a[0],x+.2,2.76,2.4,.28,{size:10,bold:true,color:C.lime});text(s,a[1],x+.2,3.18,2.42,.68,{size:16,bold:true,color:C.cream});});
 text(s,'전사·정리 과정',.78,4.7,2.4,.32,{size:12,bold:true,color:C.lime});text(s,'원문 순위 이미지 직접 확인 → 6개 그룹 × 7위 = 42개 항목 CSV화 → 각 권역 1위 추출',.78,5.15,11.65,.55,{size:16,color:C.cream});
 text(s,'구매액 원수치·패널 규모·가중 방식은 공개 페이지에서 확인하지 못했습니다.',.78,6.18,11.6,.38,{size:13,bold:true,color:C.coral});foot(s);
}
// 5 winners
{
 const s=base('THE REGIONAL LEADERS','다섯 권역 중 네 곳, 참이슬 후레쉬','공개된 권역 구분 그대로 정리했습니다. 제주도는 별도 권역으로 제시되지 않습니다.');
 const positions=[['수도권',.9,2.3],['충청권',4.25,2.3],['전라권',7.6,2.3],['강원권',.9,4.48],['경상권',7.6,4.48]];
 positions.forEach(([r,x,y])=>{const brand=winners.find(a=>a.region===r).brand;const special=brand==='진로';box(s,x,y,3.0,1.67,special?C.ink:C.white);text(s,short.get(r),x+.22,y+.18,2.5,.26,{size:11,bold:true,color:special?C.blue:C.muted2});text(s,'1위',x+.22,y+.6,.9,.65,{size:30,bold:true,color:special?C.coral:C.ink,fontFace:'Aptos'});text(s,brand,x+1.08,y+.68,1.74,.5,{size:15,bold:true,color:special?C.cream:C.ink});});
 box(s,10.95,2.3,1.65,3.85,C.lime);text(s,'4',11.2,2.75,1.1,1.0,{size:48,bold:true,color:C.ink,fontFace:'Aptos Display',align:'center'});text(s,'개 권역',11.12,3.83,1.3,.34,{size:12,bold:true,color:C.ink,align:'center'});text(s,'같은 1위',11.1,4.34,1.35,.34,{size:12,bold:true,color:C.ink,align:'center'});
 foot(s);
}
// 6 rank matrix top 5
{
 const s=base('RANKS AT A GLANCE','1위뿐 아니라 TOP 5의 배열도 닮았다','색상은 브랜드명 식별용입니다. 칸의 크기·색 농도는 구매액을 뜻하지 않습니다.',true);
 const brandsByRegion=regions.map(r=>({r, items:rows.filter(x=>x.region===r&&x.rank<=5)}));
 const x0=2.32,y0=2.38,cw=1.95,rh=.64;
 text(s,'순위',.76,y0,.95,.34,{size:11,bold:true,color:C.muted});
 brandsByRegion.forEach((g,j)=>text(s,short.get(g.r),x0+j*cw,y0,cw-.1,.34,{size:12,bold:true,color:C.cream,align:'center'}));
 for(let rank=1;rank<=5;rank++){
   text(s,String(rank),.98,y0+.48+rank*rh,.45,.28,{size:15,bold:true,color:rank===1?C.lime:C.muted,fontFace:'Aptos'});
   brandsByRegion.forEach((g,j)=>{const item=g.items.find(x=>x.rank===rank);const highlight=item.brand==='참이슬 후레쉬'?C.lime:item.brand==='진로'?C.coral:C.panel;box(s,x0+j*cw,y0+.38+rank*rh,cw-.12,.52,highlight,.08);text(s,item.brand,x0+j*cw+.05,y0+.45+rank*rh,cw-.22,.3,{size:11,bold:rank===1,color:highlight===C.panel?C.cream:C.ink,align:'center'});});
 }
 text(s,'전국 순위도 1위는 참이슬 후레쉬 · 2위 처음처럼 · 3위 진로',.78,6.45,11.8,.33,{size:13,bold:true,color:C.lime});foot(s);
}
// 7 distinctive region signal
{
 const s=base('REGIONAL SIGNALS','지역 브랜드도 순위표에 등장한다','다만 지역 권역에서 1위인지는 별개로 확인해야 합니다. 아래는 해당 권역 TOP 7 내 위치입니다.',true);
 const items=[['충청권','이제우린  #5','오투린  #6','선양  #7'],['전라권','잎새주  #5'],['경상권','좋은데이  #3','깨끗한아침 참  #4','시원  #6','대선  #7']];
 items.forEach((g,i)=>{const x=.78+i*4.18;box(s,x,2.62,3.78,2.85,C.panel);text(s,g[0],x+.24,2.92,3.2,.42,{size:18,bold:true,color:C.lime});g.slice(1).forEach((v,j)=>{let yy=3.55+j*.48;text(s,v,x+.28,yy,3.1,.34,{size:14,color:C.cream});});});
 text(s,'순위표에 보임 = 해당 권역 1위라는 뜻은 아님',.82,5.93,11.3,.44,{size:19,bold:true,color:C.coral});foot(s);
}
// 8 finding
{
 const s=base('WHAT THE DATA SAYS','지역별 1위는 다르지만, 차이는 한 권역에서 나타났다','공개된 TOP 7만으로 말할 수 있는 관찰과 해석의 경계를 나눕니다.');
 box(s,.75,2.42,5.68,3.15,'E4EBD9');text(s,'관찰',1.05,2.72,2,.36,{size:13,bold:true,color:C.muted2});text(s,'4개 권역  →  참이슬 후레쉬\n경상권      →  진로',1.05,3.37,4.9,1.1,{size:24,bold:true});text(s,'전국 1위도 참이슬 후레쉬',1.08,4.9,4.8,.32,{size:13,color:C.muted2});
 box(s,6.78,2.42,5.78,3.15,C.ink);text(s,'해석의 범위',7.08,2.72,2.5,.36,{size:13,bold:true,color:C.blue});text(s,'“가정용 구매액 순위에서\n권역별 선두 브랜드가 같다/다르다”',7.08,3.35,4.9,1.0,{size:20,bold:true,color:C.cream});text(s,'브랜드 선호 원인이나 지역 충성도를 입증한 것은 아님',7.08,4.9,4.9,.38,{size:12,color:C.muted});
 text(s,'한 시점의 순위는 차이를 보여 주지만, 왜 그런지는 설명하지 않습니다.',.82,6.15,11.5,.4,{size:15,bold:true});foot(s);
}
// 9 limitations
{
 const s=base('LIMITS & NEXT STEP','이 순위를 ‘지역별 소비량’으로 읽으면 안 된다','자료가 다루지 않거나 공개하지 않은 항목을 명시합니다.',true);
 const limits=[['구매액 수치','금액·점유율·판매량 없음'],['대상','가정용 구매만; 음식점·주점 제외'],['표본 정보','패널 규모·표집·가중 방법 미공개'],['지역 단위','5개 권역; 제주 별도 순위 없음'],['원인','가격·유통·브랜드 인지도 효과 분리 불가']];
 limits.forEach((a,i)=>{let y=2.35+i*.72;box(s,.8,y,2.1,.48,C.panel);text(s,a[0],.98,y+.09,1.7,.28,{size:12,bold:true,color:C.lime});text(s,a[1],3.2,y+.07,8.9,.32,{size:14,color:C.cream});});
 text(s,'후속 조사: 동일 기간의 지역별 판매량·구매액 원자료와 외식 채널을 함께 비교',.84,6.28,11.3,.43,{size:14,bold:true,color:C.blue});foot(s);
}
// 10 conclusion sources
{
 const s=base('CONCLUSION','가정용 구매액 1위, 네 권역의 답은 같았다','수도권·충청권·전라권·강원권은 참이슬 후레쉬, 경상권은 진로.',true);
 text(s,'이 결론은 공개된 2026년 5월 MAT TOP 7의 순위 범위 안에서만 유효합니다.',.8,2.42,11.7,.58,{size:20,bold:true,color:C.lime});
 box(s,.78,3.42,11.78,2.23,C.panel);text(s,'주요 출처',1.04,3.7,2.2,.32,{size:12,bold:true,color:C.blue});text(s,'트렌드모니터 · 엠브레인 구매딥데이터 DD.B\n“우리 동네는 이 소주 먹는데? 지역별 최애 소주 순위 TOP7” · 2026-07-02\n기준: 2026년 5월 MAT(최근 12개월) 가정용 소주류 구매액 순위',1.04,4.17,10.8,.95,{size:13,color:C.cream});
 text(s,'원문 페이지',.82,6.06,1.2,.24,{size:9,bold:true,color:C.muted});text(s,sourceUrl,.82,6.36,11.6,.26,{size:9,color:C.blue,fontFace:'Aptos'});foot(s,'전사 데이터: data/regional_top7_2026_mat.csv | 보조 해설(Puppia)은 순위 원자료 검증에 사용하지 않음');
}

await pptx.writeFile({fileName:path.join(out,'soju_regional_rank_2026.pptx')});

// Web presentation, with the same facts and an ordinal matrix. anime.js animates the rank cards.
const top7=regions.map(region=>({region,items:rows.filter(r=>r.region===region)}));
const payload=JSON.stringify({regions,winners,rows:rows.filter(r=>regions.includes(r.region)),top7,sourceUrl,sourceFigure});
const html=`<!doctype html><html lang="ko"><head><meta charset="utf-8"><meta name="viewport" content="width=device-width,initial-scale=1"><title>지역별 가정용 소주 구매액 1위는?</title><script src="https://cdn.jsdelivr.net/npm/animejs@3.2.2/lib/anime.min.js"></script><style>
:root{--bg:#101a1b;--panel:#1a2929;--cream:#f5f0e5;--muted:#a6b2aa;--lime:#c9f36a;--coral:#ff785e;--blue:#81c8d8}*{box-sizing:border-box}body{margin:0;background:#080e0f;color:var(--cream);font-family:system-ui,'Malgun Gothic',sans-serif;overflow:hidden}.stage{height:100svh;display:grid;place-items:center}.slide{position:absolute;width:min(100vw,177.78vh);height:min(56.25vw,100vh);padding:5.4% 6%;background:var(--bg);overflow:auto;opacity:0;transform:translateX(4%);pointer-events:none;transition:opacity .45s,transform .45s}.slide.active{opacity:1;transform:none;pointer-events:auto}.kicker{text-transform:uppercase;letter-spacing:.16em;color:var(--lime);font:700 clamp(10px,1vw,15px) Arial}.title{font-size:clamp(27px,3.1vw,52px);line-height:1.2;margin:.35em 0}.sub{font-size:clamp(14px,1.25vw,21px);color:var(--muted);line-height:1.55}.foot{position:absolute;bottom:3.3%;left:6%;font-size:clamp(8px,.65vw,11px);color:#718078}.hero{font-size:clamp(42px,6.2vw,96px);font-weight:850;line-height:1.08;max-width:75%;margin:9vh 0 3vh}.cover{background:radial-gradient(ellipse at 88% 46%,#314342 0,transparent 36%),var(--bg)}.cover .hero span{color:var(--lime)}.bottles{position:absolute;right:8%;top:23%;display:flex;align-items:end;gap:12px;height:43%}.bottle{width:clamp(18px,2vw,32px);height:75%;background:var(--lime);border-radius:11px 11px 5px 5px;position:relative;transform-origin:bottom}.bottle:before{content:'';position:absolute;top:-14%;left:28%;width:44%;height:17%;background:var(--cream);border-radius:4px}.bottle:nth-child(2){height:88%;background:var(--coral)}.bottle:nth-child(3){height:66%;background:var(--blue)}.bottle:nth-child(4){height:78%}.bottle:nth-child(5){height:59%;background:var(--coral)}.metric{font-size:clamp(56px,8vw,116px);font-weight:900;color:var(--lime);line-height:1}.grid2{display:grid;grid-template-columns:1fr 1fr;gap:2vw;margin-top:5vh}.card{background:var(--panel);padding:2vw;border-radius:16px;min-height:17vh}.card h3{font-size:clamp(18px,2vw,30px);margin:.3em 0}.card p{color:var(--muted);font-size:clamp(12px,1.1vw,18px)}.winners{display:grid;grid-template-columns:repeat(3,1fr);gap:1.1vw;margin-top:4vh}.winner{background:var(--panel);border-radius:14px;padding:1.3vw}.winner strong{display:block;color:var(--lime);font-size:clamp(17px,2vw,30px);margin-top:.6em}.winner.alt strong{color:var(--coral)}.matrix{display:grid;grid-template-columns:repeat(6,1fr);gap:5px;margin-top:3vh}.cell{padding:.65vw .25vw;text-align:center;border-radius:6px;background:var(--panel);font-size:clamp(9px,.88vw,15px);min-height:4vh}.cell.head{background:transparent;color:var(--muted);font-weight:700}.cell.r1{background:var(--lime);color:#101a1b;font-weight:800}.cell.jinro{background:var(--coral);color:#101a1b;font-weight:800}.rankCards{display:grid;grid-template-columns:repeat(3,1fr);gap:1vw;margin-top:4vh}.rankCards .card{min-height:28vh}.tag{display:inline-block;border:1px solid #596a66;border-radius:999px;padding:.35em .8em;margin:.3em;font-size:clamp(12px,1.2vw,19px)}.tag b{color:var(--lime)}.footerNav{position:fixed;z-index:5;bottom:13px;left:50%;transform:translateX(-50%);display:flex;align-items:center;gap:12px;background:#111d1ddd;padding:7px 13px;border:1px solid #344341;border-radius:30px}.footerNav button{background:none;border:0;color:var(--cream);font-size:18px;cursor:pointer}.progress{width:90px;height:3px;background:#344341}.progress i{display:block;height:100%;background:var(--lime);width:0}.counter{font:12px Arial;color:var(--muted)}a{color:var(--blue);overflow-wrap:anywhere}@media(max-width:700px){body{overflow:auto}.stage{height:auto;display:block}.slide{position:relative;width:100%;height:auto;min-height:100svh;padding:12vw 7vw 20vw;display:none;overflow:visible}.slide.active{display:block}.foot{position:static;margin-top:9vh}.hero{max-width:100%;margin-top:12vh}.hero span{display:block}.bottles{right:10%;top:60%;height:18%}.winners{grid-template-columns:1fr 1fr}.rankCards{grid-template-columns:1fr}.matrix{grid-template-columns:repeat(6,minmax(80px,1fr));overflow:auto}.grid2{grid-template-columns:1fr}.footerNav{bottom:9px}}
</style></head><body><main class="stage">
<section class="slide cover active"><div class="kicker">DATA / 2026 · PURCHASE RANK</div><h1 class="hero">지역별 가정용<br>소주 구매액 <span>1위는?</span></h1><p class="sub">5개 권역의 공개 TOP 7 순위로 확인한<br>우리 동네의 1위 브랜드</p><div class="bottles">${'<i class="bottle"></i>'.repeat(5)}</div><p class="sub" style="margin-top:8vh;color:var(--lime)">2026년 5월 MAT · 가정 내 구매액 기준</p><div class="foot">데이터 리터러시 프로젝트 · 2026.10</div></section>
<section class="slide"><div class="kicker">THE ANSWER</div><h2 class="title">1위는 한 브랜드에 모였을까?</h2><p class="sub">공개된 5개 권역의 가정용 소주 구매액 순위에서 확인한 결과</p><div class="grid2"><article class="card"><div class="metric">4 / 5</div><h3>권역의 1위</h3><p>수도권 · 충청권 · 전라권 · 강원권</p><span class="tag"><b>참이슬 후레쉬</b></span></article><article class="card"><div class="metric" style="color:var(--coral)">1 / 5</div><h3>다른 1위</h3><p>경상권</p><span class="tag"><b style="color:var(--coral)">진로</b></span></article></div><p class="sub" style="margin-top:3vh">‘1위’는 공개된 <b>가정용 소주류 구매액 순위</b> 기준입니다.</p><div class="foot">트렌드모니터 · 엠브레인 DD.B · 2026년 5월 MAT</div></section>
<section class="slide"><div class="kicker">RESEARCH QUESTION</div><h2 class="title">“우리 지역 소주”가 집에서도 1위일까?</h2><p class="sub">지역별 가정용 구매액 순위와 공개 TOP 7을 비교했습니다.</p><div class="grid2">${[['01','어떤 브랜드가 1위인가?','권역별 TOP 7'],['02','지역마다 달라질까?','수도권·충청·전라·경상·강원'],['03','순위가 말하지 않는 것은?','구매액 규모·점유율·외식 구매']].map(x=>`<article class="card"><div class="kicker">${x[0]}</div><h3>${x[1]}</h3><p>${x[2]}</p></article>`).join('')}</div><div class="foot">가설: 전국 유통 브랜드가 여러 권역에서 선두일 수 있다.</div></section>
<section class="slide"><div class="kicker">DATA / METHOD</div><h2 class="title">비교 대상은 ‘가정 내 구매액 순위’</h2><p class="sub">브랜드 선호도나 주류 전체 소비량과는 측정 대상이 다릅니다.</p><div class="grid2">${[['자료 제공','트렌드모니터 · 엠브레인 DD.B'],['관측 기간','2026년 5월 MAT · 최근 12개월'],['범위','전국 + 5개 권역 · 가정용'],['공개 항목','권역별 TOP 7 순위와 브랜드']].map(x=>`<article class="card"><div class="kicker">${x[0]}</div><h3>${x[1]}</h3></article>`).join('')}</div><p class="sub" style="margin-top:3vh;color:var(--coral)">구매액 원수치·패널 규모·가중 방식은 공개 페이지에서 확인할 수 없었습니다.</p></section>
<section class="slide"><div class="kicker">THE REGIONAL LEADERS</div><h2 class="title">다섯 권역 중 네 곳, 참이슬 후레쉬</h2><p class="sub">제주도는 별도 권역으로 제시되지 않습니다.</p><div class="winners">${winners.map(x=>`<article class="winner ${x.brand==='진로'?'alt':''}"><span>${short.get(x.region)}</span><strong>${x.brand}</strong><small>가정용 구매액 순위 1위</small></article>`).join('')}</div><div class="foot">출처: 트렌드모니터·엠브레인 DD.B, 2026년 5월 MAT</div></section>
<section class="slide"><div class="kicker">RANKS AT A GLANCE</div><h2 class="title">1위뿐 아니라 TOP 5 배열도 닮았다</h2><p class="sub">강조색은 브랜드 식별용이며 구매액 차이를 나타내지 않습니다.</p><div class="matrix"><div class="cell head">순위</div>${regions.map(r=>`<div class="cell head">${short.get(r)}</div>`).join('')}${[1,2,3,4,5].flatMap(rank=>[`<div class="cell head">${rank}</div>`,...regions.map(region=>{const brand=rows.find(x=>x.region===region&&x.rank===rank).brand;return `<div class="cell ${brand==='참이슬 후레쉬'?'r1':brand==='진로'?'jinro':''}">${brand}</div>`})]).join('')}</div><p class="sub" style="margin-top:3vh">전국 순위 1위 참이슬 후레쉬 · 2위 처음처럼 · 3위 진로</p></section>
<section class="slide"><div class="kicker">REGIONAL SIGNALS</div><h2 class="title">지역 브랜드도 TOP 7에 보인다</h2><p class="sub">단, 순위표에 등장하는 것과 그 지역의 1위인 것은 다릅니다.</p><div class="rankCards"><article class="card"><h3>충청권</h3><p><span class="tag">이제우린 <b>#5</b></span><span class="tag">오투린 <b>#6</b></span><span class="tag">선양 <b>#7</b></span></p></article><article class="card"><h3>전라권</h3><p><span class="tag">잎새주 <b>#5</b></span></p></article><article class="card"><h3>경상권</h3><p><span class="tag">좋은데이 <b>#3</b></span><span class="tag">깨끗한아침 참 <b>#4</b></span><span class="tag">시원 <b>#6</b></span><span class="tag">대선 <b>#7</b></span></p></article></div></section>
<section class="slide"><div class="kicker">WHAT THE DATA SAYS</div><h2 class="title">지역별 1위는 다르지만, 차이는 한 권역</h2><div class="grid2"><article class="card"><div class="kicker">관찰</div><h3>4개 권역 → 참이슬 후레쉬<br>경상권 → 진로</h3><p>전국 1위도 참이슬 후레쉬</p></article><article class="card"><div class="kicker">해석의 범위</div><h3>구매액 순위의 권역별 선두가 같다 / 다르다</h3><p>이 자료만으로 지역 충성도나 원인을 입증하지 못합니다.</p></article></div><p class="sub" style="margin-top:3vh">한 시점의 순위는 차이를 보여 주지만, 왜 그런지는 설명하지 않습니다.</p></section>
<section class="slide"><div class="kicker">LIMITS & NEXT STEP</div><h2 class="title">이 순위를 ‘지역별 소비량’으로 읽으면 안 된다</h2><div class="rankCards">${[['구매액 수치','금액·점유율·판매량 미공개'],['대상','가정용만; 음식점·주점 제외'],['표본 정보','패널·표집·가중 방식 미공개'],['지역','5개 권역; 제주 별도 순위 없음'],['원인','가격·유통·인지도 분리 불가']].map(x=>`<article class="card"><h3>${x[0]}</h3><p>${x[1]}</p></article>`).join('')}</div><p class="sub" style="margin-top:3vh">후속 조사: 지역별 실제 판매량과 외식 채널까지 비교</p></section>
<section class="slide"><div class="kicker">CONCLUSION / SOURCES</div><h2 class="title">네 권역의 가정용 구매액 1위는 같았다</h2><p class="sub" style="color:var(--lime)">수도권·충청권·전라권·강원권: 참이슬 후레쉬<br>경상권: 진로</p><article class="card" style="margin-top:4vh"><h3>주요 출처</h3><p>트렌드모니터 · 엠브레인 구매딥데이터 DD.B<br>“우리 동네는 이 소주 먹는데? 지역별 최애 소주 순위 TOP7” · 2026-07-02<br>기준: 2026년 5월 MAT, 가정용 소주류 구매액 순위</p><a href="${sourceUrl}" target="_blank" rel="noopener">원문 열기 ↗</a></article><p class="sub" style="margin-top:3vh;font-size:12px">공개된 순위만 전사해 비교했습니다. 구매액 금액은 공개되지 않았습니다.</p></section>
</main><nav class="footerNav"><button id="prev" aria-label="이전 슬라이드">←</button><span class="counter" id="counter">01 / 10</span><span class="progress"><i id="bar"></i></span><button id="next" aria-label="다음 슬라이드">→</button><button id="full" aria-label="전체 화면">⛶</button></nav><script>const slides=[...document.querySelectorAll('.slide')];let n=0;function show(k){slides[n].classList.remove('active');n=(k+slides.length)%slides.length;slides[n].classList.add('active');document.getElementById('counter').textContent=String(n+1).padStart(2,'0')+' / '+String(slides.length).padStart(2,'0');document.getElementById('bar').style.width=((n+1)/slides.length*100)+'%';if(window.anime)anime({targets:slides[n].querySelectorAll('.card,.winner,.bottle'),translateY:[18,0],opacity:[0,1],delay:anime.stagger(65),duration:650,easing:'easeOutCubic'});}document.getElementById('prev').onclick=()=>show(n-1);document.getElementById('next').onclick=()=>show(n+1);document.getElementById('full').onclick=()=>document.fullscreenElement?document.exitFullscreen():document.documentElement.requestFullscreen();document.addEventListener('keydown',e=>{if(['ArrowRight','PageDown',' '].includes(e.key)){e.preventDefault();show(n+1)}if(['ArrowLeft','PageUp'].includes(e.key)){e.preventDefault();show(n-1)}});let x=null;document.addEventListener('touchstart',e=>x=e.touches[0].clientX,{passive:true});document.addEventListener('touchend',e=>{if(x!==null){let d=e.changedTouches[0].clientX-x;if(Math.abs(d)>55)show(n+(d<0?1:-1));x=null}},{passive:true});document.getElementById('bar').style.width=(100/slides.length)+'%';</script></body></html>`;
fs.mkdirSync(path.join(root,'docs'),{recursive:true}); fs.writeFileSync(path.join(root,'docs/index.html'),html,'utf8');
fs.mkdirSync(path.join(root,'docs/data'),{recursive:true});fs.copyFileSync(path.join(root,'data/regional_top7_2026_mat.csv'),path.join(root,'docs/data/regional_top7_2026_mat.csv'));
console.log(`Wrote ${path.join(out,'soju_regional_rank_2026.pptx')} and docs/index.html (${pptx._slides.length} slides)`);
