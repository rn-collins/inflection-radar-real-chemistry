import { Presentation, PresentationFile } from "@oai/artifact-tool";
import path from "node:path";
import fs from "node:fs/promises";

const OUT = path.resolve("../..", "assets/meeting-deck.pptx");
const QA = path.resolve("../..", "qa/deck-native");

const C = {
  paper: "#F7F4EE", ink: "#14242B", teal: "#006B67", mint: "#A8E6D6",
  coral: "#F26B5E", ochre: "#D6A23A", fog: "#E8EFED", mid: "#5C6B70",
  white: "#FFFFFF",
};

const deck = Presentation.create({ slideSize: { width: 1280, height: 720 } });

function box(slide, x, y, w, h, fill, radius=0) {
  return slide.shapes.add({
    geometry: radius ? "roundRect" : "rect",
    position: { left:x, top:y, width:w, height:h },
    fill, line: { style:"solid", fill:"none", width:0 },
    ...(radius ? { borderRadius: radius } : {}),
  });
}
function text(slide, value, x, y, w, h, size=20, color=C.ink, bold=false, align="left") {
  const s=slide.shapes.add({
    geometry:"textbox", position:{left:x,top:y,width:w,height:h},
    fill:"none", line:{style:"solid",fill:"none",width:0},
  });
  s.text=value;
  s.text.style={fontFamily:"Aptos",fontSize:size,color,bold,alignment:align,verticalAlignment:"middle"};
  return s;
}
function chrome(slide, n, section="PILOT 1") {
  slide.background.fill=C.paper;
  box(slide,0,0,1280,10,C.teal);
  text(slide,"INFLECTION RADAR",64,28,300,22,12,C.teal,true);
  text(slide,section,930,28,250,22,11,C.mid,true,"right");
  box(slide,64,675,1152,1,"#C8D6D2");
  text(slide,"RAYVEN-NIKKITA COLLINS",64,686,330,18,9,C.mid,false);
  text(slide,String(n).padStart(2,"0"),1125,686,90,18,9,C.mid,true,"right");
}
function title(slide, value, sub, n, section) {
  chrome(slide,n,section);
  text(slide,value,64,78,1120,78,42,C.ink,true);
  if(sub) text(slide,sub,64,158,1060,46,19,C.mid,false);
  box(slide,64,220,120,5,C.coral);
}
function notes(slide, body, sources="") {
  slide.speakerNotes.textFrame.setText(`${body}\n\n[Sources]\n${sources || "Source: Inflection Radar internal report and public-source capability review."}`);
}

// 1
{
  const s=deck.slides.add(); s.background.fill=C.ink;
  box(s,0,0,1280,10,C.mint);
  text(s,"INFLECTION RADAR",70,52,400,30,14,C.mint,true);
  text(s,"From signal\nto growth decision",70,155,820,175,64,C.white,true);
  box(s,72,370,520,6,C.coral);
  text(s,"Pilot 1 architecture for emerging neuropsychiatry opportunity intelligence",70,405,790,70,22,C.white,false);
  box(s,922,120,210,210,C.teal,18);
  text(s,"50–75",940,158,174,68,50,C.white,true,"center");
  text(s,"companies",940,230,174,34,18,C.mint,true,"center");
  text(s,"12–15\npriority accounts",940,284,174,72,19,C.white,true,"center");
  text(s,"Prepared for Will Hargreaves / Real Chemistry\nRayven-Nikkita Collins • 28 July 2026",70,590,670,54,14,"#C8D6D2",false);
  notes(s,"Open by framing this as a decision meeting, not a product demo. The pilot is a complete external research engagement; later integration is conditional.");
}
// 2
{
  const s=deck.slides.add(); title(s,"The decision on the table","Approve Pilot 1 scope and the smallest set of inputs needed to make its portfolio useful.",2,"DECISION");
  box(s,64,270,720,260,C.white,18);
  text(s,"PROCEED NOW",94,294,250,24,13,C.teal,true);
  text(s,"Eight weeks\n50–75-company universe\n12–15 human-reviewed dossiers",94,338,620,132,31,C.ink,true);
  box(s,830,270,386,260,C.fog,18);
  text(s,"NOT REQUIRED NOW",860,294,260,24,13,C.coral,true);
  text(s,"ANATOMI integration\nCRM connection\ninternal or client data\nproduction engineering",860,340,310,138,22,C.ink,false);
  notes(s,"Ask for scope, intended reader, public service priorities, and success standard. Explicitly state what approval does not include.");
}
// 3
{
  const s=deck.slides.add(); title(s,"Will’s five questions have five direct answers","The architecture is deliberately bounded where public evidence or model output cannot resolve an internal fact.",3,"ANSWER");
  const rows=[
    ["MODEL","Bounded task portfolio; humans own inclusion, interpretation and distribution."],
    ["ASSUMPTIONS","Visible register: basis, status, owner and calibration path."],
    ["SOURCES","Permission follows every record through acquisition, use and retention."],
    ["COMPLIANCE","Company-level scope; no PHI; restrictions and escalation built in."],
    ["RELIABILITY","Evidence tiers and contradiction review prevent synthesis from becoming fact."],
  ];
  let y=260;
  for(const [a,b] of rows){
    text(s,a,66,y,180,50,14,C.teal,true);
    text(s,b,255,y,900,50,21,C.ink,false);
    box(s,64,y+56,1150,1,"#CBD6D3"); y+=72;
  }
  notes(s,"Do not unpack all mechanics here. This slide proves the diligence points were heard and answered. The next slides show the protected design decisions.");
}
// 4
{
  const s=deck.slides.add(); title(s,"Pilot 1 asks one question—and refuses to answer another","That separation prevents false precision.",4,"PILOT 1");
  box(s,64,270,500,245,C.teal,18);
  text(s,"EXTERNAL OPPORTUNITY",96,298,430,26,13,C.mint,true);
  text(s,"Which companies appear attractive for Real Chemistry review?",96,350,410,100,31,C.white,true);
  box(s,716,270,500,245,C.white,18);
  text(s,"INTERNAL PURSUIT VIABILITY",748,298,430,26,13,C.coral,true);
  text(s,"Which attractive companies are plausible internal pursuits?",748,350,410,100,31,C.ink,true);
  text(s,"Pilot 1 can build the first independently.",64,560,500,28,18,C.teal,true);
  text(s,"Only Real Chemistry can establish the second.",716,560,500,28,18,C.coral,true);
  notes(s,"Public evidence cannot establish clients, conflicts, active pursuits, relationships, economics, ownership, or practice capacity.");
}
// 5
{
  const s=deck.slides.add(); title(s,"Evidence moves through a visible reasoning chain","Every transition is inspectable; none is an automatic promotion of confidence.",5,"METHOD");
  const items=[["SOURCE","What was observed"],["FACT","What the record establishes"],["SIGNAL","What changed"],["HYPOTHESIS","Why it may matter now"],["RELEVANCE","Who should inspect it"]];
  let x=64;
  for(let i=0;i<items.length;i++){
    const [a,b]=items[i]; box(s,x,285,200,190,i===0?C.teal:C.white,16);
    text(s,String(i+1),x+18,300,38,30,16,i===0?C.mint:C.coral,true);
    text(s,a,x+18,345,165,28,16,i===0?C.white:C.teal,true);
    text(s,b,x+18,388,160,60,18,i===0?C.white:C.ink,false);
    if(i<4) text(s,"→",x+204,345,38,48,28,C.coral,true,"center");
    x+=238;
  }
  text(s,"Verified fact  •  Corroborated finding  •  Analytical inference  •  Hypothesis / early signal",64,545,1120,32,17,C.mid,false);
  notes(s,"Stress that the portfolio may include hypotheses, but they remain labeled hypotheses. The evidentiary foundation is primary sources and validated records.");
}
// 6
{
  const s=deck.slides.add(); title(s,"AI does bounded work. Humans retain authority.","Provider-flexible, but never model-indifferent.",6,"MODEL");
  const cols=[
    [C.teal,"AI-ASSISTED","Extraction\nEntity matching\nComparison\nClassification\nContradiction flags\nDraft synthesis"],
    [C.fog,"DETERMINISTIC","Source registry\nRecord IDs\nRequired fields\nWeights + date rules\nCitation links\nVersioning"],
    [C.ink,"HUMAN-ONLY","Company inclusion\nSignal meaning\nScore override\nService relevance\nEscalation\nFinal portfolio"],
  ];
  let x=64;
  for(const [fill,head,body] of cols){
    box(s,x,270,350,280,fill,18);
    text(s,head,x+28,294,294,28,15,fill===C.ink?C.mint:(fill===C.teal?C.mint:C.teal),true);
    text(s,body,x+28,345,294,165,22,fill===C.ink?C.white:C.ink,false);
    x+=400;
  }
  notes(s,"Task selection depends on sensitivity, restrictions, provider terms, retention, training use, accuracy, observability, security, cost, performance, and suitability.");
}
// 7
{
  const s=deck.slides.add(); title(s,"Permission is a path, not a source label","Restrictions travel with the record—and with intelligence derived from it.",7,"GOVERN");
  const labels=["SOURCE","METHOD","STORAGE","MODEL","OUTPUT","AUDIENCE","RETENTION"];
  let x=64;
  for(let i=0;i<labels.length;i++){
    box(s,x,300,145,92,i===0?C.teal:C.white,12);
    text(s,labels[i],x+8,322,129,28,13,i===0?C.white:C.teal,true,"center");
    if(i<6) text(s,"→",x+145,322,22,28,20,C.coral,true,"center");
    x+=164;
  }
  box(s,64,445,1152,105,C.fog,15);
  text(s,"PILOT 1 DEFAULT",92,464,200,24,13,C.teal,true);
  text(s,"Official, first-party and permitted public sources. No PHI. No unapproved scraping. No restricted licensed-source ingestion.",92,496,1060,38,21,C.ink,false);
  notes(s,"A derived field does not become unrestricted because a model produced it. Later internal work uses Real Chemistry-approved entitlements, systems, retention, and routing.");
}
// 8
{
  const s=deck.slides.add(); title(s,"The candidate white space is an operating layer","The public record supports the ingredients. It does not publicly establish this exact upstream growth workflow.",8,"FIT");
  box(s,64,268,346,270,C.white,18);
  text(s,"CONFIRMED INGREDIENTS",92,294,290,26,13,C.teal,true);
  text(s,"ANATOMI\nAnalytics & Insights\nHealthGEO + ReputAI\nMedical + communications\nAccess + media\nGrowth leadership",92,340,280,160,20,C.ink,false);
  box(s,468,268,346,270,C.fog,18);
  text(s,"NOT PUBLICLY EVIDENCED",496,294,290,26,13,C.coral,true);
  text(s,"One prospect universe\nService-need mapping\nInternal viability layer\nCross-practice routing\nOutcome learning loop",496,340,280,160,20,C.ink,false);
  box(s,872,268,344,270,C.teal,18);
  text(s,"PILOT 1 ROLE",900,294,285,26,13,C.mint,true);
  text(s,"Build the external opportunity map first.\n\nLet Real Chemistry decide whether and how to calibrate it.",900,340,270,160,23,C.white,true);
  notes(s,"Never convert 'not publicly evidenced' into 'does not exist.' Invite internal confirmation and duplication checks.");
}
// 9
{
  const s=deck.slides.add(); title(s,"Discover first. Integration remains conditional.","The method becomes more internal only after it proves useful.",9,"PROGRESSION");
  const stages=[
    ["NOW","DISCOVER","Public evidence\nExternal score\n12–15 dossiers","Which companies warrant review?",C.teal],
    ["NEXT / IF USEFUL","QUALIFY","Clients + conflicts\nOwnership + thresholds\nInternal viability","Which are plausible pursuits?",C.ochre],
    ["LATER / IF OWNED","INSTITUTIONALIZE","Monitoring + queues\nCRM + approved data\nOutcome learning","Should this become persistent?",C.coral],
  ];
  let x=64;
  for(const [tag,name,body,q,accent] of stages){
    box(s,x,270,350,300,C.white,18); box(s,x,270,350,9,accent);
    text(s,tag,x+28,298,290,22,12,accent,true);
    text(s,name,x+28,338,294,42,24,C.ink,true);
    text(s,body,x+28,400,294,82,20,C.mid,false);
    text(s,q,x+28,506,294,38,16,C.ink,true);
    x+=400;
  }
  notes(s,"Only Discover is in scope now. Qualify requires internal commercial context. Institutionalize requires ownership, approved systems, data, monitoring, and maintenance decisions.");
}
// 10
{
  const s=deck.slides.add(); title(s,"Eight weeks from scope to handoff","The pilot is large enough to expose patterns and bounded enough to evaluate.",10,"BUILD");
  const weeks=[["1","Scope"],["2","Taxonomy"],["3","Universe"],["4","Verify"],["5","Score"],["6","Dossiers"],["7","Portfolio"],["8","Handoff"]];
  let x=64;
  for(let i=0;i<weeks.length;i++){
    const [n,l]=weeks[i]; box(s,x,300,128,150,i===0?C.teal:C.white,14);
    text(s,n,x+14,315,100,50,31,i===0?C.mint:C.coral,true,"center");
    text(s,l,x+10,380,108,28,15,i===0?C.white:C.ink,true,"center");
    if(i<7) box(s,x+128,372,16,3,C.coral);
    x+=144;
  }
  box(s,64,500,1152,80,C.fog,14);
  text(s,"WEEKLY MEMO",90,522,170,22,12,C.teal,true);
  text(s,"What changed • which sources were used • what is blocked • which assumptions need review",260,513,900,38,19,C.ink,false);
  notes(s,"The final handoff includes the account portfolio, dossiers, evidence and permission records, assumption and model-task registers, rejected/deferred list, evaluation, and recommendation.");
}
// 11
{
  const s=deck.slides.add(); title(s,"A priority dossier must scan in two minutes—and survive review","The output is useful only if the reasoning and restrictions remain visible.",11,"OUTPUT");
  const items=["ACCOUNT SNAPSHOT","EVIDENCE SUMMARY","CLINICAL + REGULATORY","CORPORATE MOVEMENT","READINESS SIGNALS","NARRATIVE + REPUTATION","RC RELEVANCE","PERMISSION RECORD","HUMAN REVIEW","NEXT STEP"];
  let x=64,y=272;
  for(let i=0;i<items.length;i++){
    box(s,x,y,218,74,i===0?C.teal:C.white,12);
    text(s,String(i+1).padStart(2,"0"),x+14,y+15,34,22,12,i===0?C.mint:C.coral,true);
    text(s,items[i],x+50,y+12,150,42,13,i===0?C.white:C.ink,true);
    x+=234; if((i+1)%5===0){x=64;y+=100;}
  }
  notes(s,"The dossier should expose the source-to-decision chain, evidence tier, restrictions, contradictions, score rationale, reviewer questions, and disposition.");
}
// 12
{
  const s=deck.slides.add(); title(s,"The pilot proves usefulness—not inevitability","Success is better internal review, not a promise that every ranked company becomes an account.",12,"EVALUATE");
  const items=[
    ["RESEARCH","Citation accuracy\nCoverage\nStaleness\nContradictions"],
    ["DECISION","Usefulness\nFalse positives\nAdvance/defer rationale"],
    ["COMMERCIAL","Review-worthy accounts\nService-fit clarity\nTiming insight"],
    ["GOVERNANCE","Permissions\nRestricted-source exclusion\nAudit completeness"],
    ["LEARNING","Which signals predicted\ninterest, deferral,\npursuit or rejection"],
  ];
  let x=64;
  for(let i=0;i<items.length;i++){
    const [a,b]=items[i]; box(s,x,285,208,240,i===0?C.teal:C.white,16);
    text(s,a,x+18,306,172,26,13,i===0?C.mint:C.teal,true);
    text(s,b,x+18,355,172,120,18,i===0?C.white:C.ink,false);
    x+=236;
  }
  notes(s,"Compare the portfolio with simple baselines such as financing and clinical-stage lists and, if safely available, an internal target list. The method must add judgment, not decoration.");
}
// 13
{
  const s=deck.slides.add(); s.background.fill=C.ink; box(s,0,0,1280,10,C.mint);
  text(s,"THE DECISION",70,62,300,28,14,C.mint,true);
  text(s,"Approve Pilot 1 scope\nand its minimal inputs.",70,155,820,150,58,C.white,true);
  box(s,70,350,520,6,C.coral);
  const asks=["SCOPE","INTENDED READER","PUBLIC SERVICE PRIORITIES","SUCCESS STANDARD"];
  let y=390;
  for(const a of asks){ text(s,a,70,y,510,30,17,C.white,true); y+=50; }
  box(s,820,150,360,390,C.teal,22);
  text(s,"Reserve for later",860,190,280,32,15,C.mint,true);
  text(s,"ANATOMI\nCRM\nlicensed-source enrichment\ninternal viability scoring\nproduction integration",860,250,280,210,24,C.white,false);
  text(s,"Run Discover first.",70,620,500,35,22,C.mint,true);
  notes(s,"Close by confirming the four immediate inputs. If Will wants technical questions explored now, capture them as discovery items without making them prerequisites for the external pilot.");
}

await fs.mkdir(path.dirname(OUT), { recursive:true });
await fs.mkdir(QA, { recursive:true });
for(let i=0;i<deck.slides.items.length;i++){
  const slide=deck.slides.items[i];
  const png=await deck.export({slide,format:"png",scale:1});
  await fs.writeFile(path.join(QA,`slide-${String(i+1).padStart(2,"0")}.png`),
    Buffer.from(await png.arrayBuffer()));
}
const file=await PresentationFile.exportPptx(deck);
await file.save(OUT);
console.log(OUT);
