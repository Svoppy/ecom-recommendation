import fs from 'node:fs/promises';
import path from 'node:path';
import { pathToFileURL } from 'node:url';
import { Presentation, PresentationFile } from '@oai/artifact-tool';

const ROOT = '/Users/diaskazikhanov/Desktop/aitu/aiml/cs2';
const SKILL = '/Users/diaskazikhanov/.codex/plugins/cache/openai-primary-runtime/presentations/26.1011.11113/skills/presentations';
const NODE_MODULES = '/Users/diaskazikhanov/.cache/codex-runtimes/codex-primary-runtime/dependencies/node/node_modules';
const PYTHON = '/Users/diaskazikhanov/.cache/codex-runtimes/codex-primary-runtime/dependencies/python/bin/python3';
const BUILD = path.join(ROOT, '.build', 'pptx_build');
const FINAL = path.join(ROOT, 'deliverables', 'recommendation_presentation_final.pptx');
const FONT = 'Helvetica Neue';
const C = { bg: '#F7F9FC', ink: '#172B4D', muted: '#53657D', teal: '#087E8B', blue: '#2864A5', pale: '#E8F3F5', line: '#D9E2EC', white: '#FFFFFF', orange: '#D17B0F', red: '#A83B42', green: '#27745A' };
await fs.mkdir(BUILD, { recursive: true });
await fs.mkdir(path.dirname(FINAL), { recursive: true });
try { await fs.symlink(NODE_MODULES, path.join(BUILD, 'node_modules'), 'dir'); } catch (e) { if (e.code !== 'EEXIST') throw e; }
const resultData = JSON.parse(await fs.readFile(path.join(ROOT, 'results/model_results.json'), 'utf8'));
const results = resultData.results;
const pres = Presentation.create({ slideSize: { width: 1280, height: 720 } });

function shape(slide, geometry, x, y, w, h, fill='none', lineFill='none', lineWidth=0, name) {
  return slide.shapes.add({ geometry, name, position: { left:x, top:y, width:w, height:h }, fill, line: { style:'solid', fill:lineFill, width:lineWidth } });
}
function text(slide, value, x, y, w, h, size=22, color=C.ink, bold=false, opts={}) {
  const s = shape(slide, 'textbox', x,y,w,h,'none','none',0,opts.name);
  s.text = value;
  s.text.style = { fontSize:size, typeface:FONT, color, bold, verticalAlignment: opts.valign || 'top', wrap:true };
  return s;
}
function card(slide, x,y,w,h, title, body, opts={}) {
  shape(slide,'roundRect',x,y,w,h,opts.fill||C.white,C.line,1);
  if (opts.kicker) text(slide,opts.kicker,x+20,y+16,w-40,24,13,opts.accent||C.teal,true);
  text(slide,title,x+20,y+(opts.kicker?44:18),w-40,opts.titleH||40,opts.titleSize||21,opts.accent||C.ink,true);
  text(slide,body,x+20,y+(opts.kicker?90:66),w-40,h-(opts.kicker?102:78),opts.bodySize||17,C.muted,false);
}
function base(titleText, section, n, note) {
  const s = pres.slides.add(); s.background.fill=C.bg;
  shape(s,'rect',0,0,12,720,C.teal);
  text(s,section.toUpperCase(),64,30,440,20,12,C.teal,true);
  text(s,titleText,64,64,1152,54,34,C.ink,true);
  shape(s,'rect',64,666,1152,1,C.line);
  text(s,'Kazakhstan e-commerce • implicit-feedback recommendation benchmark',64,681,900,18,11,C.muted);
  text(s,String(n).padStart(2,'0'),1170,680,44,20,12,C.muted,true);
  if (note) s.speakerNotes.text=note;
  return s;
}
function bullets(slide, items, x=84, y=158, w=1080, size=22, gap=70) {
  items.forEach((item,i)=>{
    shape(slide,'ellipse',x,y+i*gap+8,12,12,C.teal);
    text(slide,item,x+28,y+i*gap,w-28,gap-8,size,C.ink,false);
  });
}
function metric(slide, x,y,w, value, label, accent=C.teal) {
  shape(slide,'roundRect',x,y,w,106,C.white,C.line,1);
  text(slide,value,x+16,y+18,w-32,36,30,accent,true);
  text(slide,label,x+16,y+61,w-32,34,15,C.muted,false);
}

// 1. Cover
{
 const s=pres.slides.add(); s.background.fill=C.bg;
 shape(s,'rect',0,0,18,720,C.teal); shape(s,'rect',875,0,405,720,C.pale);
 text(s,'ASSIGNMENT 1  /  MACHINE LEARNING',74,72,650,26,15,C.teal,true);
 text(s,'Regression models for\ne-commerce product\nrecommendation',74,132,760,224,47,C.ink,true);
 text(s,'A Kazakhstan-related electronics purchase-history benchmark',78,389,720,70,24,C.muted);
 shape(s,'roundRect',78,516,652,72,C.white,C.line,1);
 text(s,'11 models     •     14 features     •     10 user-grouped folds',100,539,610,32,19,C.teal,true);
 text(s,'Prepared for Assignment 1  |  2026',78,636,630,24,15,C.muted);
 s.speakerNotes.text='Study based on the Kaggle electronics-store purchase-history dataset. Source: https://www.kaggle.com/datasets/mkechinov/ecommerce-purchase-history-from-electronics-store';
}
// 2. Motivation
{
 const s=base('Why recommendations matter here','Problem framing',2,'Kazakhstan’s official statistics report ecommerce growth and continued expansion of online retail. Sources: https://stat.gov.kz/en/industries/labor-and-income/stat-ags/publications/159348/ ; https://www.gov.kz/org/mti/press/news/details/1277488?lang=en');
 text(s,'Research problem',80,150,420,28,15,C.teal,true);
 text(s,'How well can common regression models rank products a user is likely to purchase from historical implicit signals?',80,188,495,145,29,C.ink,true);
 card(s,650,150,530,154,'Context','Online retail growth increases the value of relevant product discovery. A useful model should score candidate items from sparse behavioural histories.',{kicker:'KAZAKHSTAN'});
 card(s,650,327,530,154,'Research question','Across eleven required estimators, which yields the best held-out prediction and top-10 ranking under a user-grouped protocol?',{kicker:'BENCHMARK'});
 text(s,'This is a model comparison, not a deployed end-to-end recommender.',80,500,1020,42,20,C.orange,true);
}
// 3. Literature
{
 const s=base('What recent recommender research suggests','Literature review',3,'Selected sources: Necula & Păvăloaia (2023), https://doi.org/10.3390/app13095531 ; Zangerle & Bauer (2022), https://doi.org/10.1145/3556536 ; Chen et al. (2023), https://doi.org/10.1145/3522672 ; Hidasi & Czapp (2023), https://doi.org/10.1145/3604915.3608839');
 card(s,78,160,350,350,'Implicit feedback','Purchases indicate preference, but missing interactions do not prove dislike. Negative sampling is a practical assumption that must be reported.',{kicker:'DATA',bodySize:19});
 card(s,465,160,350,350,'Ranking matters','Recommendation quality depends on which relevant items appear near the top. RMSE alone does not measure the ordering users see.',{kicker:'EVALUATION',bodySize:19});
 card(s,852,160,350,350,'Baseline discipline','Comparisons need common splits and careful protocol reporting. A benchmark result is conditional on data, features, and tuning.',{kicker:'METHOD',bodySize:19});
 text(s,'Focused search: 2021–2026 papers on recommender evaluation, implicit feedback, and e-commerce ML; not a registered systematic review.',80,557,1110,54,16,C.muted);
}
// 4. Dataset
{
 const s=base('Dataset: large event log, sparse usable identities','Materials and data',4,'Dataset: Kechinov, eCommerce purchase history from electronics store, Kaggle. https://www.kaggle.com/datasets/mkechinov/ecommerce-purchase-history-from-electronics-store');
 metric(s,80,154,250,'2.63M','raw event rows'); metric(s,356,154,250,'511,878','valid user-item-price-time rows'); metric(s,632,154,250,'121,536','historical users'); metric(s,908,154,250,'15,678','historical products');
 card(s,80,294,530,232,'Time split','History: 1 Apr–31 Aug 2020\nCutoff: 1 Sep 2020\nLater events: Sep–Nov 2020\nHistory rows: 269,975\nLater rows: 241,903',{kicker:'TEMPORAL',bodySize:18});
 card(s,640,294,530,232,'Data quality','2,069,352 raw rows lack a user ID. Price and category fields also contain missing values. Invalid timestamps and unusable rows are excluded.',{kicker:'CLEANING',bodySize:18,accent:C.orange});
 text(s,'Regional relevance is suggested by the source extract; the dataset does not document city or price currency.',80,565,1100,36,17,C.muted);
}
// 5. Target and candidates
{
 const s=base('Turning purchases into a learning target','Materials and methods',5,'Target and negative sampling are derived from the Kaggle transaction file. Source: https://www.kaggle.com/datasets/mkechinov/ecommerce-purchase-history-from-electronics-store');
 const stages=[['1','Observed later purchase','label = 1'],['2','Sample unseen user–item pair','label = 0 (assumed)'],['3','Score candidates','regression output'],['4','Sort per user','top-10 ranking']];
 stages.forEach((a,i)=>{const x=82+i*292;shape(s,'roundRect',x,213,258,190,C.white,C.line,1);shape(s,'ellipse',x+20,234,38,38,C.teal);text(s,a[0],x+32,241,20,24,19,C.white,true);text(s,a[1],x+20,295,218,54,20,C.ink,true);text(s,a[2],x+20,356,218,28,16,C.muted);if(i<3)text(s,'→',x+264,279,28,40,29,C.teal,true);});
 text(s,'6,000 positive user–item pairs • 29,318 total candidate rows • up to 4 sampled negatives per positive • seed 42',82,467,1080,46,19,C.ink,true);
 shape(s,'roundRect',82,536,1080,70,'#FFF4E4','#F0D4A5',1);
 text(s,'Interpretation: “not observed in the later window” is treated as negative for this benchmark; it is not verified user dislike.',104,557,1034,36,18,C.orange,true);
}
// 6. Features and leakage
{
 const s=base('Feature design uses only pre-cutoff history','Materials and methods',6,'Feature engineering uses events strictly before 2020-09-01; candidate target-period outcomes are excluded from feature construction. Dataset source: https://www.kaggle.com/datasets/mkechinov/ecommerce-purchase-history-from-electronics-store');
 card(s,78,158,344,365,'User history','Event count\nDistinct products\nMean log price\nDays since last event\nUser share by category\nUser share by brand',{kicker:'USER SIGNALS',bodySize:19});
 card(s,467,158,344,365,'Item history','Event count\nDistinct users\nItem share by category\nItem share by brand\nCategory ID and code\nBrand and log price',{kicker:'ITEM SIGNALS',bodySize:19});
 card(s,856,158,344,365,'Protocol guardrails','Global time cutoff before outcomes\nWhole users held out per fold\nSame sampled candidates for every model\n14 input features; one binary target',{kicker:'LEAKAGE CONTROL',bodySize:18,accent:C.green});
 text(s,'The test estimates transfer to held-out users with earlier interaction histories—not new-user cold start.',80,563,1110,40,18,C.muted);
}
// 7. Algorithms and validation
{
 const s=base('Eleven regressors, evaluated on shared folds','Methods',7,'Algorithm implementations: scikit-learn, XGBoost, LightGBM, and CatBoost. Python package versions and reproducible parameters are recorded in requirements.txt and recommendation_benchmark.py.');
 const groups=[['Linear regularization','Ridge • Lasso • Elastic Net'],['Instance-based','KNN Regression'],['Tree ensembles','Extra Trees • AdaBoost • Gradient Boosting'],['Boosting libraries','XGBoost • LightGBM • CatBoost • HistGradientBoosting']];
 groups.forEach((g,i)=>card(s,80+(i%2)*554,154+Math.floor(i/2)*176,520,145,g[0],g[1],{kicker:`MODEL FAMILY 0${i+1}`,bodySize:20}));
 shape(s,'roundRect',80,527,1074,77,C.pale,C.line,1);
 text(s,'10-fold GroupKFold by user   •   seed 42   •   RMSE + R²   •   NDCG@10 + HitRate@10',103,552,1025,36,19,C.teal,true);
}
// 8. RMSE chart
{
 const s=base('RMSE: boosted libraries lead this configuration','Results',8,'Mean RMSE over 10 user-grouped folds; lower is better. Error bars are not displayed here; standard deviations by fold are provided in results/model_results.csv.');
 const sorted=[...results].sort((a,b)=>a.rmse_mean-b.rmse_mean);
 text(s,'10-fold mean RMSE • lower is better • displayed axis: 0.32–0.36',83,139,780,26,16,C.muted,true);
 const x0=435, scale=16500, y0=180, step=34;
 sorted.forEach((r,i)=>{
   const yy=y0+i*step, width=Math.max(4,(r.rmse_mean-0.32)*scale);
   text(s,r.algorithm,88,yy,330,25,15,C.ink,i<3);
   shape(s,'roundRect',x0,yy+2,width,18,i<3?C.teal:'#9CB3C6');
   text(s,r.rmse_mean.toFixed(4),x0+width+10,yy,100,24,14,i<3?C.teal:C.muted,i<3);
 });
 shape(s,'rect',x0,y0+results.length*step+7,660,1,C.line);
 text(s,'0.32',x0-12,y0+results.length*step+14,60,21,12,C.muted);
 text(s,'0.36',x0+610,y0+results.length*step+14,60,21,12,C.muted);
 text(s,'LightGBM 0.3349  |  XGBoost 0.3350  |  CatBoost 0.3359',87,596,1060,32,18,C.ink,true);
}
// 9. Full result summary
{
 const s=base('The best model depends on the metric','Results',9,'All values are 10-fold means from results/model_results.json. NDCG@10 and HitRate@10 use sampled candidates within each held-out user.');
 const order=[...results].sort((a,b)=>a.rmse_mean-b.rmse_mean);
 const headers=['Model','RMSE ± SD','R²','NDCG@10'];
 const x=[86,556,772,956], widths=[450,200,160,190];
 shape(s,'roundRect',78,145,1105,42,C.ink);
 headers.forEach((h,i)=>text(s,h,x[i],157,widths[i],22,15,C.white,true));
 order.forEach((r,i)=>{
   const yy=193+i*37;
   if(i%2===0)shape(s,'rect',78,yy-4,1105,34,C.white);
   text(s,r.algorithm,x[0],yy,widths[0],24,15,i<3?C.teal:C.ink,i<3);
   text(s,`${r.rmse_mean.toFixed(4)} ± ${r.rmse_std.toFixed(4)}`,x[1],yy,widths[1],24,15,C.ink);
   text(s,r.r2_mean.toFixed(3),x[2],yy,widths[2],24,15,C.ink);
   text(s,r.ndcg_at_10_mean.toFixed(3),x[3],yy,widths[3],24,15,C.ink);
 });
 text(s,'LightGBM: lowest RMSE / highest R²  •  Gradient Boosting: highest NDCG@10 (0.747)',82,615,1090,28,16,C.teal,true);
}
// 10. Interpretation
{
 const s=base('Error minimization and ranking are not identical','Results and discussion',10,'Computed results: results/model_results.json and results/model_results.csv. Metrics were computed using the study’s sampled candidate set.');
 metric(s,80,161,310,'0.3349','lowest mean RMSE: LightGBM'); metric(s,426,161,310,'0.2444','highest mean R²: LightGBM'); metric(s,772,161,310,'0.7467','highest NDCG@10: Gradient Boosting');
 card(s,80,320,516,218,'Near tie','XGBoost differs from LightGBM by only 0.00002 in mean RMSE. This benchmark does not test whether that gap is statistically meaningful.',{kicker:'PREDICTION ERROR',bodySize:19});
 card(s,632,320,516,218,'Different objective','Gradient Boosting ranks first on NDCG@10 despite higher RMSE. Top-N ordering and per-candidate probability-like fit reward different behaviour.',{kicker:'RECOMMENDATION QUALITY',bodySize:19,accent:C.blue});
 text(s,'Treat the ranking scores as offline signals—not evidence of customer impact or production readiness.',82,571,1090,40,18,C.orange,true);
}
// 11. Limitations
{
 const s=base('Limits on what this benchmark can claim','Validity and limitations',11,'Source data: Kaggle dataset linked on slides 1 and 4. Protocol and caveats are documented in the manuscript and README.');
 bullets(s,[
  'Negative labels are sampled unseen pairs; item exposure is unknown.',
  'Most raw events lack user IDs; retained users may not represent all shoppers.',
  'The transaction extract is from 2020 and has no availability or impression logs.',
  'No extensive hyperparameter tuning, significance test, online A/B test, or external dataset validation.',
  'The source does not document price currency or confirm cohort-level uniqueness.'
 ],82,157,1090,19,82);
}
// 12. Conclusion and next step
{
 const s=base('Conclusion and next steps','Conclusion',12,'Study conclusions are based on the fixed benchmark run in results/model_results.json.');
 text(s,'LightGBM and XGBoost produced the strongest RMSE results; Gradient Boosting achieved the highest NDCG@10.',82,155,1060,90,28,C.ink,true);
 card(s,82,290,318,222,'Improve labels','Use impression logs and explicit exposure so “not clicked / not purchased” is better defined.',{kicker:'DATA'});
 card(s,453,290,318,222,'Broaden evaluation','Test temporal rolling windows, a fuller candidate set, and tuned baselines.',{kicker:'PROTOCOL'});
 card(s,824,290,318,222,'Check transfer','Validate on a newer Kazakhstan dataset and investigate cold-start behaviour.',{kicker:'GENERALIZATION'});
 shape(s,'roundRect',82,555,1060,57,'#FFF4E4','#F0D4A5',1);
 text(s,'Before submission: instructor must confirm cohort dataset uniqueness and applicable AI policy.',101,573,1020,26,17,C.orange,true);
}
// 13. References 1
{
 const s=base('Selected references','References',13,'Full bibliography of 15 sources is included in the research paper.');
 const refs=[
  '[1] Kazakhstan BNS (2024), E-commerce in Kazakhstan (2023); Ministry of Trade (2026), E-commerce market update.',
  '[2] Kechinov. eCommerce purchase history from electronics store. Kaggle dataset.',
  '[3] Necula & Păvăloaia (2023). AI-driven recommendations: systematic review. Applied Sciences, 13(9), 5531.',
  '[4] Zangerle & Bauer (2022). Evaluating recommender systems: survey and framework. ACM Computing Surveys, 55(8).',
  '[5] Jain & Jindal (2023). Sampling and noise filtering in recommender systems. Engineering Applications of AI, 122.',
  '[6] Yalcin & Bilge (2022). Evaluating popularity bias. Information Processing & Management, 59(6).',
  '[7] Chen et al. (2023). Revisiting negative sampling vs. non-sampling. ACM TOIS, 41(1).',
  '[8] Hidasi & Czapp (2023). Widespread flaws in offline evaluation. RecSys, pp. 848–855.'
 ];
 refs.forEach((r,i)=>text(s,r,82,144+i*57,1090,48,16,C.ink));
}
// 14. References and transparency
{
 const s=base('References and research transparency','References',14,'Full list in research_paper.docx; links and access details are recorded there.');
 const refs=[
  '[9] Zhao et al. (2023). Appropriate offline evaluation for top-N recommendation. ACM TOIS, 41(2).',
  '[10] Gusak et al. (2025). Time to split: strategies for offline evaluation. RecSys, pp. 874–883.',
  '[11] Klimashevskaia et al. (2024). Survey on popularity bias. User Modeling and User-Adapted Interaction, 34.',
  '[12] Ekstrand et al. (2024). Distributionally-informed recommender system evaluation. ACM TORS, 2(1).',
  '[13] Chen & Guestrin (2016). XGBoost: A scalable tree boosting system. KDD, pp. 785–794.',
  '[14] Ke et al. (2017). LightGBM: Efficient gradient boosting decision tree. NeurIPS, 30.',
  '[15] Prokhorenkova et al. (2018). CatBoost: Unbiased boosting with categorical features. NeurIPS, 31.'
 ];
 refs.forEach((r,i)=>text(s,r,82,141+i*48,1100,39,15,C.ink));
 shape(s,'roundRect',82,510,1080,103,C.pale,C.line,1);
 text(s,'Transparency',102,526,210,25,17,C.teal,true);
 text(s,'AI assistance was used in preparing this draft, code, and slides. No similarity score was measured. Verify citations, dataset licensing and cohort uniqueness, and follow the course AI policy.',102,557,1020,43,15,C.ink);
}

// Render every slide for visual QA and save the editable draft.
for (const [i,s] of pres.slides.items.entries()) {
  const png=await pres.export({slide:s,format:'png',scale:1});
  await fs.writeFile(path.join(BUILD,`slide-${String(i+1).padStart(2,'0')}.png`),new Uint8Array(await png.arrayBuffer()));
}
const candidate=path.join(BUILD,'candidate.pptx');
await (await PresentationFile.exportPptx(pres)).save(candidate);
const { finalizePresentation } = await import(pathToFileURL(path.join(SKILL,'container_tools/artifact_tool_utils.mjs')).href);
const result=await finalizePresentation({
  explicitTotalSlideCount:14,
  requiredNativeChartOwnerSlides:[],
  workspaceDir:ROOT,
  candidatePath:candidate,
  finalPath:FINAL,
  pythonExecutable:PYTHON,
  integrityValidatorPath:path.join(SKILL,'container_tools/inspect_presentation_package_integrity.py'),
  layoutValidatorPath:path.join(SKILL,'container_tools/inspect_presentation_layout_geometry.py'),
  layoutArgs:['--expected-slide-size-emu','12192000,6858000','--validate-bullet-geometry','--validate-heading-fit'],
  requiredNativeChartOwnerSlides:[],
  fontPolicy:{basis:'design',families:[FONT]},
  verifyArtifactToolImport:true,
  receiptPath:path.join(BUILD,'validation_final.json'),
});
console.log(JSON.stringify({finalPath:FINAL,result},null,2));
