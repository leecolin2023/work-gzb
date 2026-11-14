import json,sys
from pathlib import Path
OUT=Path(__file__).resolve().parent
p={'model':'formula_graph','scope':'江门佰富20261114现报告的算术复核，金额单位万元；并非历史样例或替代估值','inputs':{},'nodes':[],'result':'recovery','scenarios':[]}
def i(k,v,u,e):p['inputs'][k]={'value':v,'unit':u,'evidence':e}
def n(k,f,u,ref,reported=None,tol=None,**kw):
 d={'id':k,'formula':f,'unit':u,'evidence':ref,**kw}
 if reported is not None:d.update(reported=reported,tolerance=tol)
 p['nodes'].append(d)
i('principal','10930.120357','万元',['DD:P0028=109301203.57元','BK:0-概览!B7；原拟转让清单未直接读取'])
i('interest','2651.906625','万元',['DD:P0028=26519066.25元','BK:0-概览!B8'])
i('other','8.303','万元',['VR:P0193','DD:P0028','BK:0-概览!B9'])
i('priority','0','万元',['VR:P0151~P0164；按本报告取值验算，非确认股权真实净值为零'])
i('debtor_operating','3','分',['VR:T05-R03-C03','VR:P0174'])
i('debtor_assets','0','分',['VR:P0174；资产项保留0分'])
i('debtor_info_adjust','0','分',['VR:T05-R07-C04'])
i('debtor_registration_adjust','0','分',['VR:T05-R06-C04'])
i('debtor_defect_adjust','0','分',['VR:T05-R08-C04'])
i('score_floor','1','分',['VR:P0173'])
for j,(k,w,e) in enumerate([('.60','.15','T06-R05-C04/T06-R02-C05'),('.60','.15','T06-R12-C04/T06-R08-C05'),('.80','.15','T06-R18-C04/T06-R17-C05'),('.60','.20','T06-R26-C04/T06-R24-C05'),('.80','.10','T06-R31-C04/T06-R30-C05'),('.90','.10','T06-R37-C04/T06-R36-C05'),('.60','.15','T06-R45-C04/T06-R43-C05')],1):
 i('k'+str(j),k,'比例',['VR:'+e]);i('w'+str(j),w,'比例',['VR:'+e])
i('guarantor_operating','3','分',['VR:P0185'])
i('guarantor_assets','5','分',['VR:P0185；此项依据充分性待核'])
i('guarantor_adjustment','0.72','比例',['VR:T07-R02-C03；分项表未列，不能独立验证其组成'])
i('fee','0.03','比例',['VR:P0197；仅复算机构取值，未获费率论证'])
n('claim','principal + interest + other','万元',['VR:T03','DD:P0028','BK:0-概览!B6:B9'],'13590.33','.005')
n('remaining','claim - priority - other','万元',['VR:P0179~P0180'])
n('debtor_basic','max(score_floor, debtor_operating + debtor_assets + debtor_info_adjust + debtor_registration_adjust + debtor_defect_adjust) / 100','比例',['VR:T05/P0173~P0175'],'.03','.000001')
n('weight_sum','w1 + w2 + w3 + w4 + w5 + w6 + w7','比例',['VR:T06-R48-C05'],'1','.000001')
for j,row in enumerate([2,8,17,24,30,36,43],1):
 n('factor'+str(j),'k'+str(j)+' * w'+str(j),'比例',['VR:T06-R%02d-C06'%row],['.09','.09','.12','.12','.08','.09','.09'][j-1],'.000001')
n('debtor_adjustment','factor1 + factor2 + factor3 + factor4 + factor5 + factor6 + factor7','比例',['VR:T06-R48-C06'],'.68','.000001')
n('debtor_rate','debtor_basic * debtor_adjustment','比例',['VR:P0177'],'.0204','.000001')
n('debtor','remaining * debtor_rate','万元',['VR:P0179~P0181'],'277.07','.005',reported_formula_text='（债权金额－优先受偿债权价值－已单独计入的其他相关赔偿）×一般债权受偿率',reported_substitution_text='（13,590.33－0.00－8.30）×2.04%＝277.07')
n('guarantor_base','remaining - debtor','万元',['VR:P0186~P0187'],'13304.95','.005',reported_substitution_text='13,582.03－277.07＝13,304.95（展示数相减得13,304.96，未取整基数复算可解释）')
n('guarantor_basic','(guarantor_operating + guarantor_assets) / 100','比例',['VR:P0185/T07-R02-C02'],'.08','.000001')
n('guarantor_rate','guarantor_basic * guarantor_adjustment','比例',['VR:T07-R02-C04'],'.0576','.000001')
n('guarantor','guarantor_base * guarantor_rate','万元',['VR:P0188/T07-R02-C05'],'766.37','.005')
n('recovery','priority + (debtor + guarantor + other) * (1 - fee)','万元',['VR:P0196~P0198'],'1020.19','.005',reported_formula_text='优先受偿＋（债务人＋保证人＋其他赔偿）×（1－处置费率）',reported_substitution_text='0.00＋（277.07＋766.37＋8.30）×（1－3%）')
n('principal_recovery','recovery / principal','比例',['VR:P0199~P0201'],'.0933','.00005')
n('claim_recovery','recovery / claim','比例',['VR:P0202~P0204'],'.0751','.00005')
for name,key,value,why in [
 ('债务人信息不完备调整项0→-1分（条件压力）','debtor_info_adjust','-1','沿报告T05-R07的扣分项演示；不是认定必须扣分，需机构解释为何0'),
 ('保证人不动产加分5→0分（条件压力）','guarantor_assets','0','尚未核实权利负担后净值时展示该5分贡献；不判真实资产为零'),
 ('保证人调整系数72%→62%（压力）','guarantor_adjustment','.62','系数明细缺失，下降10个百分点仅作依赖展示，不是市场建议值'),
 ('处置费率3%→4%（压力）','fee','.04','增加1个百分点仅作依赖展示，不代表实际费用'),
 ('债权总额增加1000万元（信息冲突条件演示）','principal','11930.120357','BK:1-核心字段表!E27转述银行重组本金冲突；保持息费不变仅展示传导，不采认为正确余额')]:
 p['scenarios'].append({'name':name,'reason':why,'overrides':{key:{'value':value,'basis':'stress','evidence':[why]}}})
(OUT/'params_07.json').write_text(json.dumps(p,ensure_ascii=False,indent=2)+'\n',encoding='utf-8')
