"""集計のみを保存。実行: python analyze_profile.py 元CSVのパス。必要: pandas scipy statsmodels openpyxl python-docx reportlab"""
from pathlib import Path
import argparse,json,hashlib,platform
import importlib.metadata as package_metadata
import numpy as np
import pandas as pd
from scipy.stats import chi2_contingency,MonteCarloMethod
import statsmodels.api as sm
from statsmodels.stats.multitest import multipletests
from openpyxl.styles import Font,PatternFill,Alignment

p=argparse.ArgumentParser();p.add_argument('source',type=Path);args=p.parse_args()
out=Path(__file__).parent
raw=pd.read_csv(args.source,encoding='utf-8-sig');cols=list(raw.columns)
# 列の取り違えを防ぐ。
assert raw.shape==(674,19) and cols[13].startswith('Q10-1') and cols[14].startswith('Q10-2')
df=raw.rename(columns={cols[0]:'年代',cols[1]:'性別',cols[2]:'職業',cols[7]:'Q6',cols[9]:'Q7',cols[10]:'Q8',cols[13]:'参加意向'}).copy()
assert df.entry_id.nunique()==len(df)
ageorder=['10代','20代','30代','40代','50代','60代','70代','80代以上']
age_map={a:('39歳以下' if i<=2 else '70歳以上' if i>=6 else a) for i,a in enumerate(ageorder)}
df['年代群']=df['年代'].map(age_map)
ages=['39歳以下','40代','50代','60代','70歳以上']
df['層']=df['参加意向'].map({'ぜひ参加したい':'積極層','内容によっては参加したい':'潜在層','あまり参加したいと思わない':'消極層','参加したくない':'消極層'})
groups=['積極層','潜在層','消極層'];target=df['層'].eq('潜在層')
def flags(s):return s.fillna('').map(lambda v:','.join(sorted(set(x.strip() for x in v.split(',') if x.strip())))).str.get_dummies(sep=',')
blocks={k:flags(raw.iloc[:,i]) for k,i in [('Q4_実施行動',3),('Q5_難しい行動',5),('Q9_情報源',11),('Q10_企画',14)]}
actions=[x for x in blocks['Q4_実施行動'].columns if not x.startswith('その他')]
assert len(actions)==10
df['実施行動数']=blocks['Q4_実施行動'][actions].sum(axis=1)
df['難しい行動数']=blocks['Q5_難しい行動'][actions].sum(axis=1)
# Monte Carlo は疎な表に使う。固定周辺度数・固定乱数種で再現可能。
def test(ct):
 a=np.asarray(ct);a=a[a.sum(axis=1)>0];a=a[:,a.sum(axis=0)>0]
 if min(a.shape)<2:return dict(p値=np.nan,Cramers_V=np.nan,手法='変動なし',期待度数5未満割合=np.nan)
 stat,p,dof,exp=chi2_contingency(a,correction=False)
 sparse=(exp<5).mean()>.2 or exp.min()<1
 if sparse:
  res=chi2_contingency(a,correction=False,method=MonteCarloMethod(n_resamples=49999,rng=np.random.default_rng(20261030)))
  p=res.pvalue
 return dict(p値=float(p),Cramers_V=float(np.sqrt(stat/(a.sum()*min(a.shape[0]-1,a.shape[1]-1)))),手法='Monte Carlo χ²（49999回）' if sparse else 'Pearson χ²',期待度数5未満割合=float((exp<5).mean()))
def adjust(t):
 t=t.copy();mask=t['p値'].notna();t['BH補正q値']=np.nan;t.loc[mask,'BH補正q値']=multipletests(t.loc[mask,'p値'],method='fdr_bh')[1]
 t['補正後5%']=t['BH補正q値']<.05;t['補正後1%']=t['BH補正q値']<.01
 return t
# 多数の項目を同時に比較するため、主解析3群比較をまとめてBH補正。
records=[];summaries={}
for label,f in blocks.items():
 rows=[]
 for item in f:
  entry={'項目':item}
  for g in groups:
   m=df['層'].eq(g);entry[g+'人数']=int(f.loc[m,item].sum());entry[g+'分母']=int(m.sum());entry[g+'率%']=round(f.loc[m,item].mean()*100,2)
  entry['潜在−積極_ポイント']=round(entry['潜在層率%']-entry['積極層率%'],2)
  entry['潜在−消極_ポイント']=round(entry['潜在層率%']-entry['消極層率%'],2)
  rows.append(entry);records.append({'設問':label,'項目':item,**test(pd.crosstab(df['層'],f[item]))})
 summaries[label]=pd.DataFrame(rows).sort_values('潜在層率%',ascending=False)
for col in ['Q6','Q7','Q8']:
 rows=[]
 for value in sorted(df[col].unique()):
  entry={'項目':value}
  for g in groups:
   m=df['層'].eq(g);entry[g+'人数']=int((df.loc[m,col]==value).sum());entry[g+'分母']=int(m.sum());entry[g+'率%']=round((df.loc[m,col]==value).mean()*100,2)
  rows.append(entry)
 summaries[col]=pd.DataFrame(rows).sort_values('潜在層率%',ascending=False)
 records.append({'設問':col,'項目':'設問全体',**test(pd.crosstab(df['層'],df[col]))})
for col in ['年代','性別','職業']:
 rows=[]
 for val in sorted(df[col].unique()):
  e={'項目':val}
  for g in groups:
   m=df['層'].eq(g);e[g+'人数']=int(df.loc[m,col].eq(val).sum());e[g+'率%']=round(df.loc[m,col].eq(val).mean()*100,2)
  rows.append(e)
 summaries[col]=pd.DataFrame(rows)
 # 年代の細分類は参考表示のみ、検定は少数カテゴリを統合。
 c='年代群' if col=='年代' else col
 records.append({'設問':col,'項目':'設問全体',**test(pd.crosstab(df['層'],df[c]))})
comparison=adjust(pd.DataFrame(records))
# 年代差: 全回答者と潜在層内を別ファミリーで補正する。
age_tests={};age_tables={}
for scope,mask in [('全体',pd.Series(True,index=df.index)),('潜在層内',target)]:
 rr=[]
 for label,f in blocks.items():
  for item in f:
   rates=[];entry={'設問':label,'項目':item}
   for a in ages:
    m=mask & df['年代群'].eq(a);v=f.loc[m,item].mean()*100;rates.append(v);entry[f'{a}_n']=int(m.sum());entry[f'{a}_選択率%']=round(v,2)
   rr.append({'設問':label,'項目':item,**test(pd.crosstab(df.loc[mask,'年代群'],f.loc[mask,item])),'最大最小差_ポイント':max(rates)-min(rates)})
   age_tables.setdefault(scope,[]).append(entry)
 for col in ['Q6','Q7','Q8']:
  rr.append({'設問':col,'項目':'設問全体',**test(pd.crosstab(df.loc[mask,'年代群'],df.loc[mask,col]))})
  for item in sorted(df[col].unique()):
   entry={'設問':col,'項目':item}
   for a in ages:
    m=mask&df['年代群'].eq(a);entry[f'{a}_n']=int(m.sum());entry[f'{a}_選択率%']=round(df.loc[m,col].eq(item).mean()*100,2)
   age_tables[scope].append(entry)
 age_tests[scope]=adjust(pd.DataFrame(rr))
# 8年代を統合しない潜在層内の感度分析（年代統合による結果への影響を確認）。
sens=[]
for label,f in blocks.items():
 for item in f:
  sens.append({'設問':label,'項目':item,**test(pd.crosstab(df.loc[target,'年代'],f.loc[target,item]))})
for col in ['Q6','Q7','Q8']:
 sens.append({'設問':col,'項目':'設問全体',**test(pd.crosstab(df.loc[target,'年代'],df.loc[target,col]))})
age8=adjust(pd.DataFrame(sens))
# 年代群ごとの参加意向の違いは独立した1つの検定。
age_intent=pd.crosstab(df['年代群'],df['層']).reindex(ages)[groups]
age_intent_test=test(age_intent)
# 回帰: 潜在層に当てはまる関連。年齢はカテゴリで調整。性別無回答等20人は主モデルから除外。
regrows=[];modelnotes=[]
for name,other in [('潜在対それ以外',None),('潜在対積極','積極層'),('潜在対消極','消極層')]:
 mask=df['性別'].isin(['女性','男性'])
 if other:mask &= df['層'].isin(['潜在層',other])
 d=df.loc[mask].copy();x=pd.DataFrame(index=d.index)
 for a in ages[1:]:x['年代_'+a+'対39歳以下']=d['年代群'].eq(a).astype(int)
 x['男性対女性']=d['性別'].eq('男性').astype(int)
 x['実施行動数_1項目増']=d['実施行動数'];x['難しい行動数_1項目増']=d['難しい行動数']
 x['経済的理由対その他理由']=d['Q6'].str.contains('経済的理由').astype(int)
 x=sm.add_constant(x.astype(float));y=d['層'].eq('潜在層').astype(int)
 model=sm.Logit(y,x).fit(disp=False,maxiter=200);ci=model.conf_int()
 for term in model.params.index:
  if term=='const':continue
  regrows.append({'モデル':name,'説明変数':term,'対象n':len(d),'潜在層n':int(y.sum()),'オッズ比':float(np.exp(model.params[term])),'95%CI下限':float(np.exp(ci.loc[term,0])),'95%CI上限':float(np.exp(ci.loc[term,1])),'p値':float(model.pvalues[term])})
 from statsmodels.stats.outliers_influence import variance_inflation_factor
 maxvif=max(variance_inflation_factor(x.values,k) for k in range(1,x.shape[1]))
 R=np.zeros((4,len(model.params)))
 for k in range(4):R[k,k+1]=1
 age_global=model.wald_test(R,scalar=True)
 modelnotes.append({'年代の一括Wald_p':float(age_global.pvalue),'モデル':name,'n':len(d),'収束':bool(model.mle_retvals['converged']),'最大VIF':float(maxvif),'説明': '年代・性別・実施数・困難数・経済的理由を同時調整。職業は未調整。因果推論・予測精度の検証ではない。'})
reg=adjust(pd.DataFrame(regrows))
# 自由記述の原文・識別子は出力しない。
countrows=[]
for g in groups:
 d=df.loc[df['層'].eq(g)]
 for col in ['実施行動数','難しい行動数']:
  countrows.append({'層':g,'指標':col,'n':len(d),'平均':d[col].mean(),'中央値':d[col].median(),'SD':d[col].std()})
notes=[
('使用データ',args.source.name),('有効回答',len(df)),('個人識別', 'entry_idは674件すべて一意。個人単位の重複回答でないことまでは保証しない。'),
('層定義','潜在=内容によっては参加したい472、積極=ぜひ参加したい85、消極=あまり参加したいと思わない101＋参加したくない16。環境関心の高低とは別。'),
('年代統合','39歳以下（10・20・30代）、40代、50代、60代、70歳以上。若年セル不足への対応であり、若者と30代が同じとの仮定ではない。'),
('検定の帰無仮説','年代と回答は独立。p<.05なら独立を棄却し、差がある証拠。p>=.05でも差がない証明ではない。'),
('補正','層比較は全項目を一括、年代差は全体・潜在層内をそれぞれ一括、回帰は3モデル全係数を一括でBH補正。qはFDR調整値。探索的解析。'),
('Q9注意','設問は一番よく参考にする情報源だが複数回答がある。最多の情報源を1つ選んだ結果とは扱わず、各選択肢の選択率として集計。'),
('Q10-2注意','原文は3項目選択。3つ未満・超過も除外せず保持。重要結果は3つちょうどの回答に限定した参考集計も収録。'),
('Q5解釈','難しい行動の種類を問う。費用・時間・手間などの理由を直接測っていない。'),
('自由記述','個人情報保護と恣意的解釈を避け、本解析は選択式のみ。自由記述・entry_id・時刻は成果物に掲載しない。'),
('回帰','性別が男性・女性の654人。性別その他・回答しない20人を除外。女性基準。モデル間で比較対象が違う。オッズ比は確率の倍率ではない。'),
('限界','アプリ登録者等の便宜的回答。県民の母集団を無作為抽出した標本ではない。p値や信頼区間だけで代表性は回復しない。')]
audit=[]
for label,f in blocks.items():
 cnt=f.sum(axis=1);audit.append({'設問':label,'最小選択数':int(cnt.min()),'最大選択数':int(cnt.max()),'複数選択人数':int((cnt>1).sum()),'3つ未満':int((cnt<3).sum()),'3つちょうど':int((cnt==3).sum()),'3つ超過':int((cnt>3).sum())})
# 3つちょうどのQ10-2参考集計
f=blocks['Q10_企画'];m=target & f.sum(axis=1).eq(3)
ref=pd.DataFrame({'項目':f.columns,'選択人数':f.loc[m].sum().values,'選択率%':f.loc[m].mean().values*100,'分母':int(m.sum())}).sort_values('選択率%',ascending=False)
overlap=[]
for item in actions:
 a=blocks['Q4_実施行動'].loc[target,item].eq(1);b=blocks['Q5_難しい行動'].loc[target,item].eq(1)
 overlap.append({'行動':item,'分母':int(target.sum()),'実施あり_困難あり':int((a&b).sum()),'実施あり_困難なし':int((a&~b).sum()),'実施なし_困難あり':int((~a&b).sum()),'実施なし_困難なし':int((~a&~b).sum())})
sheets={'説明と注意点':pd.DataFrame(notes,columns=['項目','内容']),'データ品質':pd.DataFrame(audit),'層構成':df['層'].value_counts().rename_axis('層').reset_index(name='人数'),'年代と参加意向':age_intent.reset_index(),'年代と参加意向_検定':pd.DataFrame([age_intent_test]),'層比較_検定':comparison,'年代差_全体検定':age_tests['全体'],'年代差_潜在層検定':age_tests['潜在層内'],'参考_元8年代の検定':age8,'年代別_全体選択率':pd.DataFrame(age_tables['全体']),'年代別_潜在層選択率':pd.DataFrame(age_tables['潜在層内']),'行動数と困難数':pd.DataFrame(countrows),'実施と困難_潜在層':pd.DataFrame(overlap),'回帰_調整済み':reg,'回帰_モデル説明':pd.DataFrame(modelnotes),'参考_企画3つちょうど':ref,**summaries}
with pd.ExcelWriter(out/'潜在層の特徴_統計分析.xlsx',engine='openpyxl') as wr:
 for name,t in sheets.items():
  t.to_excel(wr,sheet_name=name,index=False);ws=wr.sheets[name];ws.freeze_panes='B2';ws.auto_filter.ref=ws.dimensions
  for c in ws[1]:c.font=Font(bold=True,color='FFFFFF');c.fill=PatternFill('solid',fgColor='24685F')
  for j,c in enumerate(t.columns,1):
   from openpyxl.utils import get_column_letter
   ws.column_dimensions[get_column_letter(j)].width=70 if c in ['項目','説明','内容','説明変数'] else 23
  for row in ws.iter_rows(min_row=2):
   for c in row:c.alignment=Alignment(wrap_text=True,vertical='top')
  for row in range(2,len(t)+2):ws.row_dimensions[row].height=48 if name=='説明と注意点' else 34
for name in ['層比較_検定','年代差_潜在層検定','回帰_調整済み']:
 sheets[name].to_csv(out/(name+'.csv'),index=False,encoding='utf-8-sig')
metadata={'source_filename':args.source.name,'sha256':hashlib.sha256(args.source.read_bytes()).hexdigest(),'n':len(df),'age_intent_test':age_intent_test,'regression':modelnotes,'python':platform.python_version(),'random_seed':20261030,'monte_carlo_repetitions':49999}
metadata['packages']={n:package_metadata.version(n) for n in ['pandas','numpy','scipy','statsmodels','patsy','openpyxl','python-docx','reportlab','matplotlib']}
(out/'解析メタデータ.json').write_text(json.dumps(metadata,ensure_ascii=False,indent=2),encoding='utf-8')
print('年代別参加意向\n',age_intent,age_intent_test)
print('\n層比較有意項目\n',comparison.loc[comparison['補正後5%'],['設問','項目','p値','BH補正q値','Cramers_V']].to_string(index=False))
for scope,t in age_tests.items():print('\n年代有意',scope,'\n',t.loc[t['補正後5%'],['設問','項目','p値','BH補正q値','Cramers_V']].to_string(index=False))
print('\n回帰\n',reg.to_string(index=False));print('\n数\n',pd.DataFrame(countrows).to_string(index=False))
