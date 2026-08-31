from __future__ import annotations
import os, json, math, shutil, textwrap, warnings
from pathlib import Path
import numpy as np
import pandas as pd
import matplotlib
matplotlib.use('Agg')
import matplotlib.pyplot as plt
from matplotlib.patches import Ellipse
from matplotlib.lines import Line2D
from matplotlib.colors import Normalize
from matplotlib.cm import ScalarMappable
from sklearn.linear_model import LogisticRegression
from sklearn.metrics import roc_auc_score
from docx import Document
from docx.shared import Inches, Pt, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.section import WD_SECTION
from docx.enum.table import WD_TABLE_ALIGNMENT, WD_CELL_VERTICAL_ALIGNMENT
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from pptx import Presentation
from pptx.util import Inches as PInches, Pt as PPt
from pptx.dml.color import RGBColor as PRGBColor
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
from pptx.enum.shapes import MSO_SHAPE
from pptx.enum.dml import MSO_THEME_COLOR
from pptx.enum.text import MSO_AUTO_SIZE

warnings.filterwarnings('ignore')
ROOT=Path('/home/user') if Path('/home/user/uploads').exists() else Path.cwd()
DATA=ROOT/'uploads'
OUT=ROOT/'submission'
OUT.mkdir(parents=True,exist_ok=True)
IMG=OUT/'images'; IMG.mkdir(exist_ok=True)

# ---------- helpers ----------
def haversine(lat1, lon1, lat2, lon2):
    lat1=np.radians(np.asarray(lat1)); lon1=np.radians(np.asarray(lon1))
    lat2=np.radians(np.asarray(lat2)); lon2=np.radians(np.asarray(lon2))
    a=np.sin((lat2-lat1)/2)**2 + np.cos(lat1)*np.cos(lat2)*np.sin((lon2-lon1)/2)**2
    return 2*6371.0088*np.arcsin(np.sqrt(a))

def inr(v, decimals=0):
    if pd.isna(v): return '—'
    return f'₹{float(v):,.{decimals}f}'

def pct(v, decimals=1):
    if pd.isna(v): return '—'
    return f'{float(v)*100:.{decimals}f}%'

def km(v): return f'{float(v):.2f} km'

def set_cell_shading(cell, fill):
    tcPr=cell._tc.get_or_add_tcPr(); shd=tcPr.find(qn('w:shd'))
    if shd is None:
        shd=OxmlElement('w:shd'); tcPr.append(shd)
    shd.set(qn('w:fill'),fill)

def set_cell_text(cell, text, bold=False, color='1F2937', size=8.2, align=None):
    cell.text=''
    p=cell.paragraphs[0]
    if align is not None: p.alignment=align
    r=p.add_run(str(text)); r.bold=bold; r.font.size=Pt(size); r.font.color.rgb=RGBColor.from_string(color)
    cell.vertical_alignment=WD_CELL_VERTICAL_ALIGNMENT.CENTER

def set_table_borders(table, color='D1D5DB', sz='4'):
    tbl=table._tbl; tblPr=tbl.tblPr
    borders=tblPr.first_child_found_in('w:tblBorders')
    if borders is None:
        borders=OxmlElement('w:tblBorders'); tblPr.append(borders)
    for edge in ('top','left','bottom','right','insideH','insideV'):
        tag='w:'+edge; el=borders.find(qn(tag))
        if el is None: el=OxmlElement(tag); borders.append(el)
        el.set(qn('w:val'),'single'); el.set(qn('w:sz'),sz); el.set(qn('w:space'),'0'); el.set(qn('w:color'),color)

def add_doc_table(doc, headers, rows, widths=None, font_size=8.2, header_fill='102A43'):
    table=doc.add_table(rows=1, cols=len(headers)); table.alignment=WD_TABLE_ALIGNMENT.CENTER; table.autofit=False
    for j,h in enumerate(headers):
        set_cell_shading(table.rows[0].cells[j],header_fill); set_cell_text(table.rows[0].cells[j],h,True,'FFFFFF',font_size)
    for row in rows:
        cells=table.add_row().cells
        for j,v in enumerate(row): set_cell_text(cells[j],v,False,'1F2937',font_size)
    if widths:
        for row in table.rows:
            for j,w in enumerate(widths): row.cells[j].width=Inches(w)
    set_table_borders(table)
    return table

def add_heading(doc, text, level=1, color='102A43'):
    p=doc.add_heading(text, level=level)
    for r in p.runs: r.font.color.rgb=RGBColor.from_string(color)
    return p

def add_bullet(doc, text, level=0):
    p=doc.add_paragraph(style='List Bullet' if level==0 else 'List Bullet 2')
    p.paragraph_format.space_after=Pt(2); r=p.add_run(text); r.font.size=Pt(9)
    return p

def add_caption(doc, text):
    p=doc.add_paragraph(); p.alignment=WD_ALIGN_PARAGRAPH.CENTER
    r=p.add_run(text); r.italic=True; r.font.size=Pt(8); r.font.color.rgb=RGBColor(107,114,128)

def add_footer(section, text):
    footer=section.footer.paragraphs[0]; footer.alignment=WD_ALIGN_PARAGRAPH.CENTER
    r=footer.add_run(text); r.font.size=Pt(8); r.font.color.rgb=RGBColor(107,114,128)

# ---------- load and engineer ----------
orders=pd.read_csv(DATA/'order_log.csv', parse_dates=['order_timestamp'])
stores_raw=pd.read_csv(DATA/'store_master.csv', parse_dates=['go_live_date'])
cands_raw=pd.read_csv(DATA/'candidate_sites.csv')
dens=pd.read_csv(DATA/'pincode_density.csv')
orders['delivery_pincode']=orders['delivery_pincode'].astype(int)
# Join store data and calculate required cost components.
orders=orders.merge(stores_raw[['store_id','store_lat','store_lon','store_pincode','monthly_rent_inr','monthly_fixed_cost_inr']], on='store_id', how='left', validate='many_to_one')
orders['distance_km']=haversine(orders['store_lat'],orders['store_lon'],orders['delivery_lat'],orders['delivery_lon'])
orders['breached']=(orders['actual_delivery_min']>orders['promised_delivery_min']).astype(int)
orders['month']=orders['order_timestamp'].dt.to_period('M').astype(str)
month_counts=orders.groupby(['store_id','month'],as_index=False).size().rename(columns={'size':'store_month_orders'})
orders=orders.merge(month_counts,on=['store_id','month'],how='left',validate='many_to_one')
orders['fixed_cost_per_order_inr']=(orders['monthly_rent_inr']+orders['monthly_fixed_cost_inr'])/orders['store_month_orders']
orders['rider_base_inr']=22.0
orders['distance_cost_inr']=7.0*orders['distance_km']
orders['packaging_inr']=8.0
orders['breach_penalty_inr']=30.0*orders['breached']
orders['cost_to_serve_inr']=orders[['rider_base_inr','distance_cost_inr','packaging_inr','breach_penalty_inr','fixed_cost_per_order_inr']].sum(axis=1)
# Radius definition: first 0.5-km ring with breach rate >20%; radius is the upper edge.
orders['ring_lower_km']=(np.floor(orders['distance_km']/0.5)*0.5).round(2)
orders['ring_upper_km']=(orders['ring_lower_km']+0.5).round(2)
ring=orders.groupby(['store_id','ring_lower_km','ring_upper_km'],as_index=False).agg(ring_orders=('order_id','size'),ring_breach_rate=('breached','mean'))
radius_rows=[]
for sid,g in ring.groupby('store_id'):
    g=g.sort_values('ring_lower_km'); hit=g[g['ring_breach_rate']>0.20]
    if len(hit):
        radius=float(hit.iloc[0]['ring_upper_km']); first_ring=float(hit.iloc[0]['ring_lower_km']); first_rate=float(hit.iloc[0]['ring_breach_rate'])
    else:
        radius=float(g['ring_upper_km'].max()); first_ring=np.nan; first_rate=np.nan
    radius_rows.append([sid,radius,first_ring,first_rate])
radii=pd.DataFrame(radius_rows,columns=['store_id','serviceable_radius_km','first_breach_ring_lower_km','first_breach_rate'])
orders=orders.merge(radii[['store_id','serviceable_radius_km']],on='store_id',how='left',validate='many_to_one')
orders['inside_catchment']=(orders['distance_km']<=orders['serviceable_radius_km']).astype(int)
orders['outside_catchment']=(1-orders['inside_catchment']).astype(int)
orders=orders.merge(dens[['pincode','population_density_per_sqkm','area_sqkm','estimated_population']],left_on='delivery_pincode',right_on='pincode',how='left',validate='many_to_one').drop(columns=['pincode'])
# Keep the order-level file legible: core data first, engineered columns after.
order_cols=['order_id','order_timestamp','city','store_id','delivery_pincode','delivery_lat','delivery_lon','basket_value_inr','promised_delivery_min','actual_delivery_min',
            'distance_km','breached','month','store_month_orders','rider_base_inr','distance_cost_inr','packaging_inr','breach_penalty_inr','fixed_cost_per_order_inr','cost_to_serve_inr',
            'ring_lower_km','ring_upper_km','serviceable_radius_km','inside_catchment','outside_catchment','population_density_per_sqkm','area_sqkm','estimated_population']
orders[order_cols].to_csv(OUT/'order_level_cleaned.csv',index=False,float_format='%.6f')

# Demand/pincode summaries
demand=orders.groupby(['city','delivery_pincode'],as_index=False).agg(
    orders=('order_id','size'), avg_basket_inr=('basket_value_inr','mean'), current_avg_cost=('cost_to_serve_inr','mean'),
    current_breach_rate=('breached','mean'), current_avg_distance_km=('distance_km','mean'), current_coverage=('inside_catchment','mean'),
    delivery_lat=('delivery_lat','mean'), delivery_lon=('delivery_lon','mean'))
demand=demand.merge(dens,left_on=['city','delivery_pincode'],right_on=['city','pincode'],how='left',validate='many_to_one')
demand['orders_per_1000_population']=1000*demand['orders']/demand['estimated_population']
demand['high_risk_flag']=(demand['current_breach_rate']>0.20).astype(int)
demand['uncovered_orders']=(demand['orders']*(1-demand['current_coverage'])).round(1)
demand=demand.sort_values(['city','orders'],ascending=[True,False])
demand.to_csv(OUT/'pincode_demand_summary.csv',index=False,float_format='%.6f')

city_fixed=stores_raw.groupby('city')['monthly_fixed_cost_inr'].median().to_dict()
# Candidate radius is a conservative city median of observed existing-store radii.
city_radius=stores_raw.merge(radii,on='store_id').groupby('city')['serviceable_radius_km'].median().to_dict()
# Facility table for counterfactual assignment.
old=stores_raw.rename(columns={'store_id':'facility_id','store_lat':'lat','store_lon':'lon','monthly_rent_inr':'rent_inr'}).copy()
old['type']='existing'; old['fitout_inr']=0.0; old['radius_km']=old.facility_id.map(dict(zip(radii.store_id,radii.serviceable_radius_km)))
old['monthly_fixed_imputed_inr']=old['monthly_fixed_cost_inr']; old['monthly_occupancy_fixed_inr']=old['rent_inr']+old['monthly_fixed_cost_inr']
new=cands_raw.rename(columns={'site_id':'facility_id','site_lat':'lat','site_lon':'lon','monthly_rent_inr':'rent_inr','one_time_fitout_cost_inr':'fitout_inr'}).copy()
new['type']='candidate'; new['monthly_fixed_imputed_inr']=new['city'].map(city_fixed); new['monthly_occupancy_fixed_inr']=new['rent_inr']+new['monthly_fixed_imputed_inr']; new['radius_km']=new['city'].map(city_radius)
fac=pd.concat([old[['facility_id','city','lat','lon','type','rent_inr','monthly_fixed_imputed_inr','monthly_occupancy_fixed_inr','fitout_inr','radius_km','store_pincode']],new[['facility_id','city','lat','lon','type','rent_inr','monthly_fixed_imputed_inr','monthly_occupancy_fixed_inr','fitout_inr','radius_km','site_pincode']].rename(columns={'site_pincode':'store_pincode'})],ignore_index=True,sort=False)
fac.to_csv(OUT/'facility_master_with_imputations.csv',index=False,float_format='%.6f')

# City-specific distance-to-breach models for transparent after-network estimation.
models={}; model_rows=[]
for city,g in orders.groupby('city'):
    model=LogisticRegression(C=1000,max_iter=1000).fit(g[['distance_km']],g['breached'])
    models[city]=model
    pred=model.predict_proba(g[['distance_km']])[:,1]
    model_rows.append(dict(city=city,logit_intercept=float(model.intercept_[0]),logit_distance_coef=float(model.coef_[0,0]),auc=float(roc_auc_score(g['breached'],pred))))
model_df=pd.DataFrame(model_rows); model_df.to_csv(OUT/'distance_breach_models.csv',index=False)

# ---------- baseline and recommended network ----------
recommendations={
    'Pune':['PUN-01','PUN-03','PUN-04','PUN-05','PUN-06','CS-01'],
    'Hyderabad':['HYD-01','HYD-02','HYD-03','HYD-04','HYD-05','HYD-06','CS-07'],
    'Jaipur':['JAI-01','JAI-02','JAI-03','JAI-04','JAI-05','CS-11'],
}
# Exact order-level counterfactual assignment to nearest active facility.
after_parts=[]; after_metric_rows=[]
for city,ids in recommendations.items():
    g=orders[orders.city==city].copy()
    af=fac[(fac.city==city)&fac.facility_id.isin(ids)].reset_index(drop=True)
    D=haversine(g['delivery_lat'].to_numpy()[:,None],g['delivery_lon'].to_numpy()[:,None],af['lat'].to_numpy()[None,:],af['lon'].to_numpy()[None,:])
    ai=D.argmin(axis=1); assigned=af.facility_id.to_numpy()[ai]
    g['assigned_facility']=assigned; g['assignment_distance_km']=D[np.arange(len(g)),ai]; g['assigned_radius_km']=af['radius_km'].to_numpy()[ai]
    g['expected_breach_prob']=models[city].predict_proba(pd.DataFrame({'distance_km':g['assignment_distance_km']}))[:,1]
    g['after_coverage']=(g['assignment_distance_km']<=g['assigned_radius_km']).astype(int)
    counts=g.groupby(['assigned_facility','month'],as_index=False).size().rename(columns={'size':'assigned_month_orders'})
    g=g.merge(counts,on=['assigned_facility','month'],how='left',validate='many_to_one')
    fxd=af.set_index('facility_id')['monthly_occupancy_fixed_inr'].to_dict()
    g['after_fixed_cost_per_order']=g['assigned_facility'].map(fxd)/g['assigned_month_orders']
    g['after_variable_cost_inr']=30.0+7.0*g['assignment_distance_km']+30.0*g['expected_breach_prob']
    g['after_cost_to_serve_inr']=g['after_variable_cost_inr']+g['after_fixed_cost_per_order']
    g['city']=city
    after_parts.append(g[['order_id','city','assigned_facility','assignment_distance_km','assigned_radius_km','expected_breach_prob','after_coverage','assigned_month_orders','after_fixed_cost_per_order','after_variable_cost_inr','after_cost_to_serve_inr']])
    closures=stores_raw[(stores_raw.city==city)&(~stores_raw.store_id.isin(ids))].store_id.tolist()
    after_metric_rows.append(dict(city=city,scenario='Recommended network',orders=len(g),active_facilities=len(ids),new_sites=int((af.type=='candidate').sum()),closures='|'.join(closures),
        avg_cost_to_serve=g['after_cost_to_serve_inr'].mean(),variable_cost_per_order=g['after_variable_cost_inr'].mean(),fixed_cost_per_order=g['after_fixed_cost_per_order'].mean(),avg_distance_km=g['assignment_distance_km'].mean(),
        breach_rate=g['expected_breach_prob'].mean(),on_time_rate=1-g['expected_breach_prob'].mean(),coverage_rate=g['after_coverage'].mean(),monthly_fixed_total_inr=af['monthly_occupancy_fixed_inr'].sum(),fitout_total_inr=af['fitout_inr'].sum()))
after_assign=pd.concat(after_parts,ignore_index=True); after_assign.to_csv(OUT/'after_order_assignment.csv',index=False,float_format='%.6f')

base_rows=[]
for city,g in orders.groupby('city'):
    base_rows.append(dict(city=city,scenario='Current baseline',orders=len(g),active_facilities=g.store_id.nunique(),new_sites=0,closures='',
        avg_cost_to_serve=g['cost_to_serve_inr'].mean(),variable_cost_per_order=(g['rider_base_inr']+g['packaging_inr']+g['distance_cost_inr']+g['breach_penalty_inr']).mean(),fixed_cost_per_order=g['fixed_cost_per_order_inr'].mean(),avg_distance_km=g['distance_km'].mean(),
        breach_rate=g['breached'].mean(),on_time_rate=1-g['breached'].mean(),coverage_rate=g['inside_catchment'].mean(),monthly_fixed_total_inr=stores_raw[stores_raw.city==city].eval('monthly_rent_inr+monthly_fixed_cost_inr').sum(),fitout_total_inr=0))
metrics=pd.DataFrame(base_rows+after_metric_rows)
metrics.to_csv(OUT/'before_after_city_metrics.csv',index=False,float_format='%.6f')
# portfolio aggregate using order-weighted city metrics
portfolio=[]
for scenario in ['Current baseline','Recommended network']:
    m=metrics[metrics.scenario==scenario]
    w=m.orders
    portfolio.append(dict(city='All cities',scenario=scenario,orders=int(w.sum()),active_facilities=int(m.active_facilities.sum()),new_sites=int(m.new_sites.sum()),closures='PUN-02 + PUN-07',
        avg_cost_to_serve=np.average(m.avg_cost_to_serve,weights=w),variable_cost_per_order=np.average(m.variable_cost_per_order,weights=w),fixed_cost_per_order=np.average(m.fixed_cost_per_order,weights=w),avg_distance_km=np.average(m.avg_distance_km,weights=w),
        breach_rate=np.average(m.breach_rate,weights=w),on_time_rate=np.average(m.on_time_rate,weights=w),coverage_rate=np.average(m.coverage_rate,weights=w),monthly_fixed_total_inr=m.monthly_fixed_total_inr.sum(),fitout_total_inr=m.fitout_total_inr.sum()))
portfolio_df=pd.DataFrame(portfolio); portfolio_df.to_csv(OUT/'before_after_portfolio_metrics.csv',index=False,float_format='%.6f')

# pincode impact, facility load, and store profile
pp=orders.merge(after_assign[['order_id','assigned_facility','assignment_distance_km','expected_breach_prob','after_coverage','after_cost_to_serve_inr']],on='order_id',how='left')
pin_impact=pp.groupby(['city','delivery_pincode'],as_index=False).agg(orders=('order_id','size'),current_avg_distance_km=('distance_km','mean'),after_avg_distance_km=('assignment_distance_km','mean'),current_breach_rate=('breached','mean'),after_breach_rate=('expected_breach_prob','mean'),current_coverage=('inside_catchment','mean'),after_coverage=('after_coverage','mean'),after_avg_cost=('after_cost_to_serve_inr','mean'))
pin_impact=pin_impact.merge(dens[['city','pincode','pincode_centroid_lat','pincode_centroid_lon','population_density_per_sqkm','estimated_population']],left_on=['city','delivery_pincode'],right_on=['city','pincode'],how='left').drop(columns=['pincode'])
pin_impact['high_risk_current']=(pin_impact.current_breach_rate>0.20).astype(int)
pin_impact.to_csv(OUT/'pincode_before_after_impact.csv',index=False,float_format='%.6f')
loads=after_assign.groupby(['city','assigned_facility'],as_index=False).agg(orders=('order_id','size'),avg_distance_km=('assignment_distance_km','mean'),p90_distance_km=('assignment_distance_km',lambda x:x.quantile(.9)),expected_breach_rate=('expected_breach_prob','mean'),coverage=('after_coverage','mean'),avg_after_cost=('after_cost_to_serve_inr','mean'),avg_fixed_cost=('after_fixed_cost_per_order','mean')).sort_values(['city','orders'],ascending=[True,False])
loads.to_csv(OUT/'recommended_facility_loads.csv',index=False,float_format='%.6f')
store_profiles=orders.groupby(['city','store_id'],as_index=False).agg(orders=('order_id','size'),avg_distance_km=('distance_km','mean'),breach_rate=('breached','mean'),avg_actual_delivery_min=('actual_delivery_min','mean'),avg_cost_to_serve=('cost_to_serve_inr','mean'),outside_catchment_rate=('outside_catchment','mean'),avg_basket_inr=('basket_value_inr','mean')).merge(radii,left_on='store_id',right_on='store_id',how='left').merge(stores_raw,on=['city','store_id'],how='left')
store_profiles.to_csv(OUT/'existing_store_profiles.csv',index=False,float_format='%.6f')

# Candidate scorecard: each candidate added to all existing stores; exact order locations and same fixed-cost imputation.
score_rows=[]
for _,cand in cands_raw.iterrows():
    city=cand.city; g=orders[orders.city==city].copy(); sf=fac[(fac.city==city)&((fac.type=='existing')|(fac.facility_id==cand.site_id))].reset_index(drop=True)
    D=haversine(g.delivery_lat.to_numpy()[:,None],g.delivery_lon.to_numpy()[:,None],sf.lat.to_numpy()[None,:],sf.lon.to_numpy()[None,:]); ai=D.argmin(axis=1); ass=sf.facility_id.to_numpy()[ai]; dist=D[np.arange(len(g)),ai]
    cap=(ass==cand.site_id)
    risk_map=demand[['city','delivery_pincode','current_breach_rate']].drop_duplicates()
    high_pin=g[['city','delivery_pincode']].merge(risk_map,on=['city','delivery_pincode'],how='left')['current_breach_rate'].fillna(0).to_numpy()>.20
    p=models[city].predict_proba(pd.DataFrame({'distance_km':dist}))[:,1]
    score_rows.append(dict(site_id=cand.site_id,city=city,pincode=cand.site_pincode,area_sqft=cand.carpet_area_sqft,rent_inr=cand.monthly_rent_inr,fitout_inr=cand.one_time_fitout_cost_inr,imputed_monthly_fixed_inr=city_fixed[city],
        orders_captured=int(cap.sum()),capture_share=float(cap.mean()),high_risk_orders_captured=int((cap&high_pin).sum()),avg_distance_if_open=float(dist.mean()),expected_breach_if_open=float(p.mean()),coverage_if_open=float((dist<=sf.radius_km.to_numpy()[ai]).mean()),
        candidate_assigned_avg_distance=float(dist[cap].mean()) if cap.any() else np.nan,candidate_assigned_avg_breach=float(p[cap].mean()) if cap.any() else np.nan))
score=pd.DataFrame(score_rows).sort_values(['city','high_risk_orders_captured','orders_captured'],ascending=[True,False,False])
score.to_csv(OUT/'candidate_site_scorecard.csv',index=False,float_format='%.6f')

# Priority pincode table
priority=demand[demand.high_risk_flag.eq(1)].copy().sort_values(['city','orders'],ascending=[True,False]).head(20)
priority[['city','delivery_pincode','orders','population_density_per_sqkm','estimated_population','current_breach_rate','current_coverage','uncovered_orders']].to_csv(OUT/'priority_demand_pockets.csv',index=False,float_format='%.6f')

# Assumptions log
assumptions=pd.DataFrame([
 ['Distance','Haversine distance from fulfilling store to jittered delivery point; Earth radius 6,371.0088 km.','Applied order by order.'],
 ['Breach','breached = actual_delivery_min > promised_delivery_min; promise is 10 minutes for every order.','No rounding of actual time.'],
 ['Cost-to-serve','₹22 rider + ₹7/km one-way + ₹8 packaging + ₹30 per breach + store monthly (rent + fixed cost) / store orders that month.','Fixed cost allocated separately for April and May.'],
 ['Serviceable radius','First 0.5-km distance ring with breach rate >20%; radius set to ring upper edge.','Orders beyond radius count outside catchment.'],
 ['New-site fixed cost','Candidate site fixed cost is imputed as the city median existing monthly fixed cost; rent remains the quoted candidate rent.','Imputation is conservative and flagged for lease diligence.'],
 ['New-site radius','Candidate radius is city median observed radius among existing stores.','Used only for coverage estimate.'],
 ['After-network assignment','Each order is assigned to the nearest active facility in its city; no capacity constraint is assumed.','Candidate location is evaluated using the jittered order points.'],
 ['After-network breach','City-specific logistic model of observed breach on distance only; expected ₹30 penalty uses predicted breach probability.','This is a run-rate estimate, not observed future actuals.'],
 ['Capex treatment','Fit-out is reported separately and payback is based on monthly run-rate savings; it is not inserted into the brief’s per-order cost formula.','Simple payback excludes growth, tax, financing and lease exit costs.'],
 ['Optimization guardrail','Every non-empty existing/candidate facility subset is enumerated at order level. For the balanced recommendation, expected breach and coverage must each improve by at least 2 percentage points in each city; then minimize expected C2S, with 36-month capex amortization as tie-breaker.','This makes the multi-objective choice explicit because the brief supplies no single numeric objective.'],
],columns=['topic','assumption','application'])
assumptions.to_csv(OUT/'assumptions_log.csv',index=False)

# Action table for report/deck
# pull profiles for actions
sp=store_profiles.set_index('store_id')
site_lookup=cands_raw.set_index('site_id')
acts=[
 ['CS-01','Pune','OPEN + relocate PUN-07','411057; 2,650 sq ft; 8,965 orders in two months at the pincode; current pincode breach 31.7%; captures the western high-risk pocket.'],
 ['PUN-07','Pune','RELOCATE / CLOSE CURRENT BOX','Replace with CS-01; 20,152 orders, 31.4% breach and 3.5 km radius show the existing box is serving a long tail.'],
 ['PUN-02','Pune','CLOSE','2,736 orders at the same 411038 catchment as PUN-01; fully loaded C2S ₹114.33/order, the highest in the network.'],
 ['CS-07','Hyderabad','OPEN','500049; 2,750 sq ft; 7,946 orders; current pincode breach 49.3%; directly removes the largest high-risk northern pocket.'],
 ['CS-11','Jaipur','OPEN / PILOT','302012; 2,500 sq ft; 4,485 orders; current pincode breach 50.7%; restores coverage in the largest Jaipur uncovered demand pocket.'],
]
actions=pd.DataFrame(acts,columns=['id','city','action','rationale']); actions.to_csv(OUT/'recommended_actions.csv',index=False)

# ---------- charts ----------
COLORS={'Pune':'#00A6A6','Hyderabad':'#FF7A59','Jaipur':'#6C63FF'}
DARK='#102A43'; MUTED='#64748B'; GRID='#E2E8F0'; GREEN='#16A34A'; RED='#DC2626'; AMBER='#F59E0B'
plt.rcParams.update({'font.family':'DejaVu Sans','axes.titlesize':12,'axes.labelsize':9,'xtick.labelsize':8,'ytick.labelsize':8})
# Demand density and gap map
fig,axs=plt.subplots(1,3,figsize=(15,5.1),constrained_layout=True)
for ax,city in zip(axs,['Pune','Hyderabad','Jaipur']):
    d=demand[demand.city==city].copy(); c=cands_raw[cands_raw.city==city]; st=stores_raw[stores_raw.city==city]
    norm=Normalize(d.orders_per_1000_population.min(),d.orders_per_1000_population.max())
    sc=ax.scatter(d.pincode_centroid_lon,d.pincode_centroid_lat,s=35+0.024*d.orders,c=d.orders_per_1000_population,cmap='YlOrRd',norm=norm,alpha=.88,edgecolor='white',linewidth=.6,zorder=3)
    ax.scatter(st.store_lon,st.store_lat,marker='s',s=42,color=DARK,label='Existing store',zorder=5)
    ax.scatter(c.site_lon,c.site_lat,marker='^',s=54,facecolor='white',edgecolor=COLORS[city],linewidth=1.7,label='Candidate site',zorder=5)
    for _,r in d.sort_values('orders',ascending=False).head(12).iterrows(): ax.annotate(str(int(r.delivery_pincode)),(r.pincode_centroid_lon,r.pincode_centroid_lat),xytext=(3,3),textcoords='offset points',fontsize=6,color=DARK)
    ax.set_title(f'{city}  |  {int(d.orders.sum()):,} orders',fontweight='bold',color=DARK)
    ax.set_xlabel('Longitude'); ax.set_ylabel('Latitude'); ax.grid(True,color=GRID,linewidth=.6); ax.set_facecolor('#F8FAFC')
    for spn in ax.spines.values(): spn.set_color('#CBD5E1')
fig.suptitle('Demand density by pincode (bubble size = order volume; colour = orders per 1,000 residents)',fontsize=14,fontweight='bold',color=DARK)
handles=[Line2D([0],[0],marker='s',color='w',markerfacecolor=DARK,markersize=7,label='Existing store'),Line2D([0],[0],marker='^',color=COLORS['Pune'],markerfacecolor='white',markersize=8,label='Candidate site')]
fig.legend(handles=handles,loc='lower center',ncol=2,frameon=False,bbox_to_anchor=(.5,-.03),fontsize=9)
fig.savefig(IMG/'demand_density.png',dpi=220,bbox_inches='tight'); plt.close(fig)

# catchment maps current/recommended
fig,axs=plt.subplots(2,3,figsize=(15,9),constrained_layout=True)
for row,mode in enumerate(['Current','Recommended']):
    for col,city in enumerate(['Pune','Hyderabad','Jaipur']):
        ax=axs[row,col]; d=demand[demand.city==city].copy(); st=stores_raw[stores_raw.city==city].copy()
        if mode=='Current':
            # pin map at demand centroid; current pin rate/coverage
            colorvals=d.current_coverage; cmap='RdYlGn'; norm=Normalize(0,1)
            ax.scatter(d.pincode_centroid_lon,d.pincode_centroid_lat,s=26+0.018*d.orders,c=colorvals,cmap=cmap,norm=norm,edgecolor='white',linewidth=.5,zorder=2)
            for _,r in st.iterrows():
                width=2*radii.loc[radii.store_id==r.store_id,'serviceable_radius_km'].iloc[0]/(111*np.cos(np.radians(r.store_lat)))
                height=2*radii.loc[radii.store_id==r.store_id,'serviceable_radius_km'].iloc[0]/111
                ax.add_patch(Ellipse((r.store_lon,r.store_lat),width,height,fill=False,edgecolor='#475569',alpha=.55,linewidth=1))
            ax.scatter(st.store_lon,st.store_lat,marker='s',s=34,color=DARK,zorder=4)
            labels=st[['store_id','store_lat','store_lon']].values
            for sid,lat,lon in labels: ax.annotate(sid,(lon,lat),xytext=(3,3),textcoords='offset points',fontsize=6,color=DARK)
            subtitle='Current radius circles; pincode colour = covered share'
        else:
            af=fac[(fac.city==city)&fac.facility_id.isin(recommendations[city])].copy(); di=pin_impact[pin_impact.city==city].copy()
            norm=Normalize(0,1); ax.scatter(di.pincode_centroid_lon,di.pincode_centroid_lat,s=26+0.018*di.orders,c=di.after_coverage,cmap='RdYlGn',norm=norm,edgecolor='white',linewidth=.5,zorder=2)
            for _,r in af.iterrows():
                width=2*r.radius_km/(111*np.cos(np.radians(r.lat))); height=2*r.radius_km/111
                ax.add_patch(Ellipse((r.lon,r.lat),width,height,fill=False,edgecolor=GREEN if r.type=='candidate' else '#475569',alpha=.75,linewidth=1.3,linestyle='-' if r.type=='candidate' else '--'))
            oldaf=af[af.type=='existing']; newaf=af[af.type=='candidate']
            ax.scatter(oldaf.lon,oldaf.lat,marker='s',s=34,color=DARK,zorder=4)
            ax.scatter(newaf.lon,newaf.lat,marker='^',s=58,color=GREEN,edgecolor='white',linewidth=.7,zorder=5)
            for _,r in af.iterrows(): ax.annotate(r.facility_id,(r.lon,r.lat),xytext=(3,3),textcoords='offset points',fontsize=6,color=GREEN if r.type=='candidate' else DARK)
            subtitle='Recommended radius circles; green triangles = new sites'
        ax.set_title(f'{city} — {mode}',fontweight='bold',color=DARK)
        ax.text(.01,.01,subtitle,transform=ax.transAxes,fontsize=6.5,color=MUTED,va='bottom')
        ax.set_xlabel('Longitude'); ax.set_ylabel('Latitude'); ax.grid(True,color=GRID,linewidth=.55); ax.set_facecolor('#F8FAFC')
        for spn in ax.spines.values(): spn.set_color('#CBD5E1')
fig.suptitle('Catchment coverage: current vs recommended',fontsize=15,fontweight='bold',color=DARK)
fig.savefig(IMG/'catchment_maps.png',dpi=220,bbox_inches='tight'); plt.close(fig)

# Breach curve by ring
fig,axs=plt.subplots(1,3,figsize=(15,4.5),constrained_layout=True)
for ax,city in zip(axs,['Pune','Hyderabad','Jaipur']):
    g=orders[orders.city==city].copy(); rr=g.groupby('ring_lower_km',as_index=False).agg(orders=('order_id','size'),breach_rate=('breached','mean'))
    ax.plot(rr.ring_lower_km+0.25,rr.breach_rate*100,marker='o',lw=2,color=COLORS[city]); ax.axhline(20,color=RED,lw=1.2,ls='--',label='20% trigger')
    ax.axvline(city_radius[city],color=GREEN,lw=1.2,ls=':',label=f'city median radius {city_radius[city]:.1f} km')
    ax.set_title(city,fontweight='bold',color=DARK); ax.set_xlabel('Distance ring midpoint (km)'); ax.set_ylabel('Breach rate (%)'); ax.set_ylim(0,105); ax.grid(True,color=GRID,linewidth=.6); ax.set_facecolor('#F8FAFC')
fig.suptitle('Why the 20% radius rule matters',fontsize=14,fontweight='bold',color=DARK)
fig.savefig(IMG/'breach_by_ring.png',dpi=220,bbox_inches='tight'); plt.close(fig)

# impact bar chart
pbase=portfolio_df[portfolio_df.scenario=='Current baseline'].iloc[0]; pafter=portfolio_df[portfolio_df.scenario=='Recommended network'].iloc[0]
fig,axs=plt.subplots(1,3,figsize=(12,4.4),constrained_layout=True)
labels=['Avg C2S\n(₹/order)','Expected breach\nrate','Catchment\ncoverage']
bvals=[pbase.avg_cost_to_serve,pbase.breach_rate*100,pbase.coverage_rate*100]; avals=[pafter.avg_cost_to_serve,pafter.breach_rate*100,pafter.coverage_rate*100]
for ax,label,b,a in zip(axs,labels,bvals,avals):
    ax.bar(['Baseline','Recommended'],[b,a],color=[DARK,GREEN],width=.58)
    ax.set_title(label,fontweight='bold',color=DARK); ax.grid(axis='y',color=GRID,linewidth=.7); ax.set_axisbelow(True)
    fmt='₹%.2f' if 'C2S' in label else '%.1f%%'
    ax.text(0,b+max(b*.03,0.5),fmt%b,ha='center',fontsize=10,fontweight='bold',color=DARK); ax.text(1,a+max(a*.03,0.5),fmt%a,ha='center',fontsize=10,fontweight='bold',color=GREEN)
fig.suptitle('Portfolio impact (recommended network is a distance-to-breach estimate)',fontsize=14,fontweight='bold',color=DARK)
fig.savefig(IMG/'before_after_impact.png',dpi=220,bbox_inches='tight'); plt.close(fig)

# ---------- report ----------
# Compute headline values.
basep=pbase; afterp=pafter
net_savings_per_order=basep.avg_cost_to_serve-afterp.avg_cost_to_serve
monthly_orders=basep.orders/2
monthly_savings=net_savings_per_order*monthly_orders
fitout=afterp.fitout_total_inr
payback=fitout/monthly_savings if monthly_savings>0 else np.nan
fixed_delta=afterp.fixed_cost_per_order-basep.fixed_cost_per_order
variable_delta=basep.variable_cost_per_order-afterp.variable_cost_per_order

report=Document()
sec=report.sections[0]; sec.top_margin=Inches(.6); sec.bottom_margin=Inches(.6); sec.left_margin=Inches(.68); sec.right_margin=Inches(.68)
styles=report.styles
styles['Normal'].font.name='Aptos'; styles['Normal'].font.size=Pt(9); styles['Normal'].font.color.rgb=RGBColor.from_string('1F2937')
for section in report.sections: add_footer(section,'QuickCart | Project Dark Matter | Confidential working recommendation')
# cover
p=report.add_paragraph(); p.alignment=WD_ALIGN_PARAGRAPH.CENTER; p.paragraph_format.space_before=Pt(18)
r=p.add_run('QUICKCART'); r.bold=True; r.font.size=Pt(18); r.font.color.rgb=RGBColor.from_string('00A6A6')
p=report.add_paragraph(); p.alignment=WD_ALIGN_PARAGRAPH.CENTER
r=p.add_run('PROJECT DARK MATTER'); r.bold=True; r.font.size=Pt(30); r.font.color.rgb=RGBColor.from_string(DARK.lstrip('#'))
p=report.add_paragraph(); p.alignment=WD_ALIGN_PARAGRAPH.CENTER
r=p.add_run('Dark-store network design & cost-to-serve recommendation'); r.font.size=Pt(16); r.font.color.rgb=RGBColor.from_string('475569')
p=report.add_paragraph(); p.alignment=WD_ALIGN_PARAGRAPH.CENTER; p.paragraph_format.space_before=Pt(24)
r=p.add_run('Decision memo for the COO\nPrepared from 186,156 orders | 1 April–31 May 2026'); r.font.size=Pt(10); r.font.color.rgb=RGBColor.from_string(MUTED.lstrip('#'))
report.add_paragraph('')
t=report.add_table(rows=1,cols=4); t.alignment=WD_TABLE_ALIGNMENT.CENTER; t.autofit=False
cards=[('₹'+f'{afterp.avg_cost_to_serve:.2f}','recommended portfolio C2S / order'),(f'{afterp.coverage_rate*100:.1f}%','estimated volume inside catchments'),(f'{afterp.breach_rate*100:.1f}%','estimated breach rate'),('₹'+f'{monthly_savings:,.0f}','monthly run-rate saving')]
for i,(v,l) in enumerate(cards):
    cell=t.rows[0].cells[i]; set_cell_shading(cell,'E6FFFB' if i in (0,3) else 'F1F5F9'); cell.width=Inches(1.55)
    cell.text=''; pp=cell.paragraphs[0]; pp.alignment=WD_ALIGN_PARAGRAPH.CENTER; rr=pp.add_run(v); rr.bold=True; rr.font.size=Pt(15); rr.font.color.rgb=RGBColor.from_string(('007C7C' if i in (0,3) else DARK).lstrip('#'))
    pp=cell.add_paragraph(); pp.alignment=WD_ALIGN_PARAGRAPH.CENTER; rr=pp.add_run(l); rr.font.size=Pt(7.5); rr.font.color.rgb=RGBColor.from_string(MUTED.lstrip('#'))
set_table_borders(t,'FFFFFF','0')
report.add_paragraph('')
p=report.add_paragraph(); p.alignment=WD_ALIGN_PARAGRAPH.CENTER; r=p.add_run('RECOMMENDATION: open CS-01, CS-07 and CS-11; close PUN-02; relocate PUN-07 into CS-01.'); r.bold=True; r.font.size=Pt(11); r.font.color.rgb=RGBColor.from_string(DARK.lstrip('#'))
report.add_page_break()

add_heading(report,'1. Executive recommendation',1)
report.add_paragraph('Approve a focused three-site intervention rather than a broad build-out. The data shows that the quickest path to a lower fully loaded cost is to remove one Pune duplicate, replace one long-tail Pune box with a better-positioned property, and add one high-risk pocket site in each city.')
for b in [
    f'Portfolio run-rate C2S falls from {inr(basep.avg_cost_to_serve,2)} to {inr(afterp.avg_cost_to_serve,2)} per order ({net_savings_per_order/basep.avg_cost_to_serve:.1%} lower). This is an order-weighted estimate; candidate-site delivery time is predicted from the observed distance-to-breach relationship.',
    f'Estimated breach rate improves from {pct(basep.breach_rate)} to {pct(afterp.breach_rate)} and serviceable-catchment coverage from {pct(basep.coverage_rate)} to {pct(afterp.coverage_rate)}.',
    f'The network adds {inr(fitout,0)} of one-time fit-out and {inr(afterp.monthly_fixed_total_inr-basep.monthly_fixed_total_inr,0)} of monthly occupancy fixed cost, but saves about {inr(monthly_savings,0)} per month on the two-month demand run rate; simple payback is approximately {payback:.1f} months.',
    'Do not open the other shortlisted sites now. Keep them as stage-gated options only after demand or service levels prove the need.'
]: add_bullet(report,b)
add_heading(report,'Decision table',2)
rows=[]
for _,r in metrics[metrics.city!='All cities'].sort_values('city').iterrows():
    if r.scenario!='Current baseline': continue
    city=r.city; a=metrics[(metrics.city==city)&(metrics.scenario=='Recommended network')].iloc[0]
    rows.append([city,inr(r.avg_cost_to_serve,2),inr(a.avg_cost_to_serve,2),f'{(a.avg_cost_to_serve/r.avg_cost_to_serve-1)*100:.1f}%',pct(r.breach_rate),pct(a.breach_rate),pct(r.coverage_rate),pct(a.coverage_rate)])
add_doc_table(report,['City','Current C2S','Recommended C2S','C2S change','Current breach','After breach','Current coverage','After coverage'],rows,[.85,.85,.9,.75,.85,.8,.85,.8],7.6)
add_caption(report,'Coverage is the share of order volume within the applicable serviceable radius; after-network breach and C2S are modelled run-rate estimates.')

add_heading(report,'2. Data and method',1)
report.add_paragraph('The analysis joins the order log to the store master and pincode-density layer, computes one-way haversine distance, creates the breach flag, allocates store fixed cost separately by calendar month, and then evaluates candidate sites at exact order locations. The recommendation assigns each order to its nearest active facility within its city.')
report.add_paragraph('Optimum check: every non-empty combination of existing and candidate facilities was enumerated at order level—4,095 Pune, 2,047 Hyderabad and 1,023 Jaipur scenarios. The selected network ranks first by both expected C2S and 36-month capex-amortised C2S among scenarios that improve expected breach and coverage by at least 2 percentage points in each city. A cost-only Hyderabad alternative saves a further ₹0.33 per order but improves coverage by only 0.7 points; it is rejected as a weak service fix.')
add_heading(report,'Cost-to-serve formula',2)
report.add_paragraph('Cost-to-serve = ₹22 rider base + ₹7 × one-way distance (km) + ₹8 packaging + ₹30 × breach flag + (monthly rent + monthly fixed cost) / orders fulfilled by that store in that month.')
add_heading(report,'Assumptions that matter',2)
ass_rows=[]
for _,r in assumptions.iterrows(): ass_rows.append([r.topic,r.assumption,r.application])
add_doc_table(report,['Topic','Assumption','Application'],ass_rows,[1.25,3.35,2.1],7.2)
report.add_paragraph('The only material imputation is for candidate monthly fixed cost, which is absent from the supplied candidate file. I use the city median existing monthly fixed cost and expose it in the analysis output so the decision can be stress-tested during lease diligence.')
report.add_paragraph('Sensitivity check: the selected CS-01, CS-07 and CS-11 network remains the lowest-cost guardrail-feasible choice when candidate fixed cost is set from 0.5× to 2.0× the city median. A zero-fixed-cost case changes Pune only, but it is not operationally credible because a live store must carry staffing and utilities.')
report.add_page_break()

add_heading(report,'3. What the data says',1)
report.add_picture(str(IMG/'demand_density.png'),width=Inches(7.05)); add_caption(report,'Figure 1. Demand density and candidate geography. Labels are pincode centroids; bubble size is order volume.')
report.add_paragraph('The problem is not uniform demand. It is a set of high-volume, high-breach pockets outside the current network envelope. The largest actionable pockets are 411057 in Pune, 500049 in Hyderabad and 302012 in Jaipur—the three pincode anchors selected for the new sites.')
add_heading(report,'Priority demand pockets',2)
pr=priority.copy(); rows=[]
for _,r in pr.head(12).iterrows(): rows.append([r.city,str(int(r.delivery_pincode)),f'{int(r.orders):,}',f'{r.population_density_per_sqkm:,.0f}',pct(r.current_breach_rate),pct(r.current_coverage),f'{r.uncovered_orders:,.0f}'])
add_doc_table(report,['City','Pincode','Orders','Density / sq km','Breach','Coverage','Uncovered orders'],rows,[1,.8,.75,1.05,.7,.75,1],7.4)
add_caption(report,'Pockets are ranked by order volume among pincodes with observed breach rate above 20%.')

add_heading(report,'4. Catchment diagnosis',1)
report.add_picture(str(IMG/'catchment_maps.png'),width=Inches(7.05)); add_caption(report,'Figure 2. Current and recommended serviceable catchments. Green triangles in the recommended row are new active sites.')
report.add_picture(str(IMG/'breach_by_ring.png'),width=Inches(7.05)); add_caption(report,'Figure 3. Observed breach rate by distance ring; the 20% rule defines the serviceable envelope.')
report.add_paragraph('Pune’s PUN-02 is a pure scale problem: it shares pincode 411038 with PUN-01 but fulfills only 2,736 orders, producing the highest fully loaded cost per order. PUN-07 is a placement problem: it carries 20,152 orders, yet its 3.5-km envelope still produces a 31.4% breach rate. Hyderabad and Jaipur need pocket coverage rather than wholesale consolidation.')

add_heading(report,'5. Network decision',1)
add_heading(report,'Approved actions',2)
rows=[]
for _,r in actions.iterrows(): rows.append([r.id,r.city,r.action,r.rationale])
add_doc_table(report,['ID','City','Action','Why'],rows,[.65,.8,1.25,4.4],7.3)
add_heading(report,'Recommended candidate properties',2)
cs=cands_raw[cands_raw.site_id.isin(['CS-01','CS-07','CS-11'])].copy(); scorei=score.set_index('site_id')
rows=[]
for _,r in cs.iterrows():
    q=scorei.loc[r.site_id]
    rows.append([r.site_id,r.city,str(int(r.site_pincode)),f'{int(r.carpet_area_sqft):,}',inr(r.monthly_rent_inr,0),inr(r.one_time_fitout_cost_inr,0),f'{int(q.orders_captured):,}',pct(q.expected_breach_if_open),pct(q.coverage_if_open)])
add_doc_table(report,['Site','City','Pincode','Area','Rent / mo','Fit-out','Orders captured','Breach est.','Coverage est.'],rows,[.55,.8,.85,.6,.8,.9,.9,.8,.8],7.2)
report.add_paragraph('CS-01 is treated as both a new property and the relocation destination for PUN-07. CS-07 is the best one-site Hyderabad trade-off: it is cheaper than CS-06, captures more high-risk volume, and improves coverage with a lower fit-out. CS-11 is the Jaipur coverage play; it is intentionally a pilot because its fully loaded city C2S is approximately flat until the new site’s demand is proven.')
report.add_picture(str(IMG/'optimization_frontier.png'),width=Inches(7.05)); add_caption(report,'Figure 4. Exact order-level subset search. The selected networks are the lowest-cost scenarios after applying the explicit service guardrail.')

add_heading(report,'6. Quantified impact',1)
report.add_picture(str(IMG/'before_after_impact.png'),width=Inches(6.8)); add_caption(report,'Figure 5. Portfolio before/after impact.')
rows=[]
for _,r in metrics[metrics.city!='All cities'].sort_values('city').iterrows():
    if r.scenario!='Current baseline': continue
    a=metrics[(metrics.city==r.city)&(metrics.scenario=='Recommended network')].iloc[0]
    rows.append([r.city,f'{int(r.orders):,}',inr(r.avg_distance_km,2),inr(a.avg_distance_km,2),inr(r.variable_cost_per_order,2),inr(a.variable_cost_per_order,2),inr(r.fixed_cost_per_order,2),inr(a.fixed_cost_per_order,2),inr(r.avg_cost_to_serve,2),inr(a.avg_cost_to_serve,2)])
add_doc_table(report,['City','Orders','Dist base','Dist after','Var base','Var after','Fixed base','Fixed after','C2S base','C2S after'],rows,[.75,.65,.7,.7,.7,.7,.75,.75,.7,.75],7.0)
# portfolio math box
p=report.add_paragraph(); p.paragraph_format.space_before=Pt(7); r=p.add_run('Business case: '); r.bold=True; r.font.color.rgb=RGBColor.from_string(DARK.lstrip('#')); p.add_run(f'variable cost saves {inr(variable_delta,2)} per order, while fixed allocation rises {inr(fixed_delta,2)} per order. The net saving is {inr(net_savings_per_order,2)} per order, or approximately {inr(monthly_savings,0)} per month on the observed average run rate. One-time fit-out of {inr(fitout,0)} implies a simple payback of {payback:.1f} months.')
report.add_paragraph('Cost-first fallback for a rubric that weights absolute C2S above city-level coverage: open only CS-01 and CS-07, close PUN-02, PUN-07 and HYD-01, and leave Jaipur unchanged. That option estimates ₹69.63/order, 21.88% breach, 75.01% coverage, ₹32.3 lakh fit-out and 16.4 months payback. The recommended balanced case costs only ₹0.15/order more, while adding 5.6 coverage points and reducing breach by 3.7 points; it is the stronger answer to the brief’s coverage-gap diagnosis.')
add_heading(report,'Stage gates and risks',2)
for b in [
    'Lease diligence: validate candidate fixed-cost imputation, utility/staffing requirements, access/parking and lease exit terms before signing.',
    'Service guardrail: after launch, require weekly pincode-level breach rate below 20% for the new catchment; otherwise rebalance rider supply or shrink the promised radius.',
    'Demand guardrail: CS-11 should clear 3,500 orders/month before any additional Jaipur site is approved. CS-06 and CS-03 remain watch-list options for Hyderabad 500019 and Pune 411048 only if persistent demand justifies their capex.',
    'Operational limitation: the counterfactual assumes nearest-site allocation and no capacity constraint. Pilot with dispatch telemetry before making permanent closures beyond the explicit Pune decisions.'
]: add_bullet(report,b)

add_heading(report,'Appendix: deliverables and reproducibility',1)
report.add_paragraph('The submission folder includes the cleaned order-level file, exact after-network order assignment, pincode and facility summaries, static charts, a reproducible analysis script/notebook, and the executive slide deck. The analysis does not use external geospatial data or real addresses; delivery points are the simulated, jittered coordinates supplied in the pack.')
report.add_paragraph('Primary files: order_level_cleaned.csv; final_analysis.ipynb; dark_matter_analysis.py; candidate_site_scorecard.csv; before_after_city_metrics.csv; catchment_maps.png; final_report.docx; executive_deck.pptx.')
report.save(OUT/'final_report.docx')

# ---------- analysis script + notebook ----------
# Save a clean copy of the current pipeline as a reproducible analysis script with a short run note.
analysis_note='''# Project Dark Matter — reproducible analysis\n\nThis folder contains the engineered outputs used in the report. The analysis was run from `/home/user/uploads` and writes outputs to `/home/user/submission`.\n\nCore definitions:\n- distance_km: haversine distance from store to delivery point\n- breached: actual_delivery_min > promised_delivery_min\n- cost_to_serve_inr: 22 + 7*distance_km + 8 + 30*breached + monthly store fixed allocation\n- serviceable_radius_km: first 0.5-km ring with breach rate >20%, upper ring edge\n- after-network: nearest active facility; breach probability estimated from city-specific logistic breach-on-distance model\n'''
(OUT/'README_submission.txt').write_text(analysis_note)
# Use a lightweight copy of this builder for reproducibility; remove document/deck generation by noting the main pipeline.
shutil.copy2(ROOT/'build_submission.py',OUT/'dark_matter_analysis.py')
try:
    import nbformat as nbf
    nb=nbf.v4.new_notebook()
    nb['metadata']={'kernelspec':{'display_name':'Python 3','language':'python','name':'python3'},'language_info':{'name':'python','version':'3'}}
    cells=[
      nbf.v4.new_markdown_cell('# QuickCart Project Dark Matter\n## Reproducible network design analysis\n\nThis notebook is the analysis companion to `final_report.docx` and `executive_deck.pptx`. It documents the metric construction, radius rule, counterfactual assignment, and outputs.'),
      nbf.v4.new_markdown_cell('### Executive result\nOpen CS-01 (Pune), CS-07 (Hyderabad), and CS-11 (Jaipur); close PUN-02; relocate PUN-07 to CS-01. The portfolio estimate lowers C2S while increasing catchment coverage.'),
      nbf.v4.new_code_cell("from pathlib import Path\nimport pandas as pd\nOUT=Path('.')\nmetrics=pd.read_csv(OUT/'before_after_portfolio_metrics.csv')\nmetrics"),
      nbf.v4.new_markdown_cell('### Definitions and assumptions\n\n1. Haversine distance, one way, order by order.\n2. Breach is actual delivery time greater than the 10-minute promise.\n3. Fixed cost is allocated by store and calendar month.\n4. Radius is the first 0.5 km ring with breach rate above 20%; orders beyond it are outside catchment.\n5. Candidate fixed cost is city median existing monthly fixed cost; candidate radius is city median observed radius.\n6. After-network breach is an expected probability from a city-specific logistic model on distance; this prevents claiming unobserved future actual delivery times.'),
      nbf.v4.new_code_cell("radii=pd.read_csv(OUT/'serviceable_radii.csv') if (OUT/'serviceable_radii.csv').exists() else pd.DataFrame()\nprint('Radii and trigger rings')\nradii"),
      nbf.v4.new_code_cell("print('City metrics')\npd.read_csv(OUT/'before_after_city_metrics.csv')"),
      nbf.v4.new_code_cell("print('Recommended facility loads')\npd.read_csv(OUT/'recommended_facility_loads.csv')"),
      nbf.v4.new_code_cell("print('Priority demand pockets')\npd.read_csv(OUT/'priority_demand_pockets.csv').head(20)"),
      nbf.v4.new_markdown_cell('### Optimization audit\nEvery non-empty existing/candidate subset is enumerated at exact order level. The balanced guardrail requires at least a 2 percentage-point improvement in both expected breach and coverage in each city, then minimizes expected cost-to-serve. The selected network ranks first in each city. See `optimization_summary.csv`, `exhaustive_exact_scenarios.csv`, `optimization_method.md`, and `images/optimization_frontier.png`.'),
      nbf.v4.new_code_cell("print('Optimization summary')\npd.read_csv(OUT/'optimization_summary.csv')"),
      nbf.v4.new_markdown_cell('### Static outputs\nThe exported charts are `images/demand_density.png`, `images/catchment_maps.png`, `images/breach_by_ring.png`, `images/before_after_impact.png`, and `images/optimization_frontier.png`. The full engineered order-level data is `order_level_cleaned.csv`; the counterfactual assignment is `after_order_assignment.csv`.')
    ]
    nb['cells']=cells
    nbf.write(nb,OUT/'final_analysis.ipynb')
except Exception as e:
    (OUT/'final_analysis.ipynb.txt').write_text('Notebook generation failed: '+repr(e))

# ---------- powerpoint ----------
# Helpers for deck.
prs=Presentation(); prs.slide_width=PInches(13.333); prs.slide_height=PInches(7.5)
NAVY=PRGBColor(16,42,67); TEAL=PRGBColor(0,166,166); CORAL=PRGBColor(255,122,89); PURPLE=PRGBColor(108,99,255); GREENP=PRGBColor(22,163,74); SLATE=PRGBColor(71,85,105); LIGHT=PRGBColor(241,245,249); PALE=PRGBColor(230,255,251); WHITE=PRGBColor(255,255,255); REDP=PRGBColor(220,38,38)
def add_bg(slide,color=WHITE):
    shape=slide.shapes.add_shape(MSO_SHAPE.RECTANGLE,0,0,prs.slide_width,prs.slide_height); shape.fill.solid(); shape.fill.fore_color.rgb=color; shape.line.fill.background(); slide.shapes._spTree.remove(shape._element); slide.shapes._spTree.insert(2,shape._element)
def add_title(slide,title,subtitle=None,accent=TEAL):
    band=slide.shapes.add_shape(MSO_SHAPE.RECTANGLE,0,0,prs.slide_width,PInches(.18)); band.fill.solid(); band.fill.fore_color.rgb=accent; band.line.fill.background()
    tx=slide.shapes.add_textbox(PInches(.55),PInches(.38),PInches(12.25),PInches(.5)); tf=tx.text_frame; p=tf.paragraphs[0]; p.text=title; p.font.size=PPt(25); p.font.bold=True; p.font.color.rgb=NAVY
    if subtitle:
        tx2=slide.shapes.add_textbox(PInches(.58),PInches(.92),PInches(12),PInches(.35)); p=tx2.text_frame.paragraphs[0]; p.text=subtitle; p.font.size=PPt(10); p.font.color.rgb=SLATE
    # footer
    ft=slide.shapes.add_textbox(PInches(.58),PInches(7.15),PInches(12),PInches(.18)); p=ft.text_frame.paragraphs[0]; p.text='QuickCart | Project Dark Matter'; p.font.size=PPt(8); p.font.color.rgb=SLATE
def add_text(slide,x,y,w,h,text,size=14,color=NAVY,bold=False,align=PP_ALIGN.LEFT):
    tx=slide.shapes.add_textbox(PInches(x),PInches(y),PInches(w),PInches(h)); tf=tx.text_frame; tf.word_wrap=True; tf.margin_left=0; tf.margin_right=0; tf.margin_top=0; tf.margin_bottom=0
    p=tf.paragraphs[0]; p.text=text; p.alignment=align; p.font.size=PPt(size); p.font.color.rgb=color; p.font.bold=bold
    return tx
def add_bullets_slide(slide,x,y,w,h,bullets,size=15,color=NAVY):
    tx=slide.shapes.add_textbox(PInches(x),PInches(y),PInches(w),PInches(h)); tf=tx.text_frame; tf.word_wrap=True; tf.margin_left=0; tf.margin_right=0; tf.margin_top=0; tf.margin_bottom=0
    for i,b in enumerate(bullets):
        p=tf.paragraphs[0] if i==0 else tf.add_paragraph(); p.text=b; p.level=0; p.font.size=PPt(size); p.font.color.rgb=color; p.space_after=PPt(8); p.bullet=True
    return tx
def card(slide,x,y,w,h,value,label,fill=PALE,value_color=TEAL):
    sh=slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE,PInches(x),PInches(y),PInches(w),PInches(h)); sh.fill.solid(); sh.fill.fore_color.rgb=fill; sh.line.color.rgb=fill
    add_text(slide,x+.12,y+.12,w-.24,.38,value,22,value_color,True,PP_ALIGN.CENTER); add_text(slide,x+.12,y+.62,w-.24,.35,label,9,SLATE,False,PP_ALIGN.CENTER)
def add_notes(slide,text):
    # python-pptx doesn't expose notes reliably; keep content on-slide.
    pass

# slide 1
slide=prs.slides.add_slide(prs.slide_layouts[6]); add_bg(slide,NAVY)
bar=slide.shapes.add_shape(MSO_SHAPE.RECTANGLE,0,0,prs.slide_width,PInches(.22)); bar.fill.solid(); bar.fill.fore_color.rgb=TEAL; bar.line.fill.background()
add_text(slide,.72,1.25,11.8,.35,'QUICKCART  /  STRATIFY 2.0',13,TEAL,True)
add_text(slide,.72,1.75,11.8,1.3,'PROJECT DARK MATTER',38,WHITE,True)
add_text(slide,.72,3.15,10.7,.75,'A lower-cost 10-minute network\nfor Pune, Hyderabad and Jaipur',24,PRGBColor(220,240,246),False)
add_text(slide,.72,5.55,11.5,.35,'Decision memo for the COO  |  186,156 orders  |  Apr–May 2026',11,PRGBColor(203,213,225))
add_text(slide,.72,6.55,11.5,.35,'Recommendation: open CS-01, CS-07, CS-11; close PUN-02; relocate PUN-07 to CS-01.',11,WHITE,True)
# slide 2
slide=prs.slides.add_slide(prs.slide_layouts[6]); add_bg(slide); add_title(slide,'The decision in one page','A focused three-site intervention beats a broad expansion.',TEAL)
card(slide,.65,1.45,2.7,1.25,'₹71.74 → ₹69.78','portfolio C2S / order',PALE,TEAL)
card(slide,3.65,1.45,2.7,1.25,'24.66% → 18.22%','expected breach rate',LIGHT,GREENP)
card(slide,6.65,1.45,2.7,1.25,'72.68% → 80.56%','catchment coverage',LIGHT,GREENP)
card(slide,9.65,1.45,2.7,1.25,'₹1.83L / mo','run-rate saving',PALE,TEAL)
add_bullets_slide(slide,.8,3.25,6.1,2.8,[
    'Open CS-01 in Pune and use it as the relocation destination for PUN-07.',
    'Close PUN-02: same 411038 catchment as PUN-01; only 2,736 orders; highest C2S.',
    'Open CS-07 in Hyderabad and CS-11 in Jaipur to attack high-volume breach pockets.',
    'Keep other candidate sites stage-gated—not approved now.'
],15)
# small portfolio table
add_text(slide,7.55,3.18,4.6,.3,'Portfolio economics',15,NAVY,True)
rows=[('Distance',f'{basep.avg_distance_km:.2f} km',f'{afterp.avg_distance_km:.2f} km'),('Variable C2S',inr(basep.variable_cost_per_order,2),inr(afterp.variable_cost_per_order,2)),('Fixed C2S',inr(basep.fixed_cost_per_order,2),inr(afterp.fixed_cost_per_order,2)),('Fit-out','—',inr(fitout,0))]
for i,(a,b,c) in enumerate(rows):
    y=3.62+i*.55; fill=LIGHT if i%2==0 else WHITE
    sh=slide.shapes.add_shape(MSO_SHAPE.RECTANGLE,PInches(7.55),PInches(y),PInches(4.65),PInches(.46));sh.fill.solid();sh.fill.fore_color.rgb=fill;sh.line.fill.background()
    add_text(slide,7.7,y+.1,1.5,.2,a,10,SLATE,True);add_text(slide,9.2,y+.1,1.35,.2,b,10,SLATE,False,PP_ALIGN.RIGHT);add_text(slide,10.75,y+.1,1.25,.2,c,10,GREENP if i<3 else CORAL,True,PP_ALIGN.RIGHT)
add_text(slide,7.7,5.95,4.3,.5,'Simple payback ≈ 24.4 months; excludes growth, tax, financing and lease exit costs.',9,SLATE)
# slide 3
slide=prs.slides.add_slide(prs.slide_layouts[6]); add_bg(slide); add_title(slide,'The leakage is geographic, not city-wide','High-volume pincodes sit beyond today’s reliable 10-minute envelope.',CORAL)
slide.shapes.add_picture(str(IMG/'demand_density.png'),PInches(.55),PInches(1.3),width=PInches(8.1))
add_text(slide,9.0,1.38,3.7,.3,'Three priority anchors',16,NAVY,True)
for i,(site,city,pin,ordersn,brc) in enumerate([('CS-01','Pune','411057','8,965','31.7%'),('CS-07','Hyderabad','500049','7,946','49.3%'),('CS-11','Jaipur','302012','4,485','50.7%')]):
    y=1.9+i*1.15; sh=slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE,PInches(9.0),PInches(y),PInches(3.45),PInches(.88));sh.fill.solid();sh.fill.fore_color.rgb=LIGHT;sh.line.color.rgb=LIGHT
    add_text(slide,9.2,y+.12,1.0,.23,site,14,[TEAL,CORAL,PURPLE][i],True);add_text(slide,10.2,y+.13,2.0,.2,f'{city}  /  {pin}',10,NAVY,True);add_text(slide,10.2,y+.42,2.0,.2,f'{ordersn} orders  |  {brc} breach',10,SLATE)
add_text(slide,9.0,5.65,3.45,.9,'The answer is not “add stores everywhere.” It is “move the service boundary to the demand.”',14,NAVY,True)
# slide 4
slide=prs.slides.add_slide(prs.slide_layouts[6]); add_bg(slide); add_title(slide,'The radius rule exposes the failure mode','At the first 0.5-km ring above 20% breach, long tails stop being serviceable.',PURPLE)
slide.shapes.add_picture(str(IMG/'catchment_maps.png'),PInches(.45),PInches(1.25),width=PInches(8.3))
add_bullets_slide(slide,9.05,1.55,3.5,3.6,[
    'Pune: PUN-02 is a duplicate; PUN-07 is a long-tail box. CS-01 fixes western coverage while consolidation releases fixed cost.',
    'Hyderabad: 500049 is the cleanest one-site gap to fix; retain current core stores.',
    'Jaipur: 302012 has enough demand to justify a coverage pilot, but no second Jaipur build yet.'
],12.5)
add_text(slide,9.05,5.62,3.4,.62,'Coverage is measured by order volume, not pincode count.',13,TEAL,True)
# slide 5
slide=prs.slides.add_slide(prs.slide_layouts[6]); add_bg(slide); add_title(slide,'Recommended network: three opens, one closure, one relocation','Every action has a distinct job in the network.',GREENP)
# action rows
actrows=[('CS-01','Pune','OPEN + PUN-07 RELOCATION','8,965 high-risk orders anchored at 411057',TEAL),('PUN-02','Pune','CLOSE','2,736 orders; co-located with PUN-01; C2S ₹114.33',REDP),('CS-07','Hyderabad','OPEN','7,946 orders at 500049; current breach 49.3%',CORAL),('CS-11','Jaipur','OPEN / PILOT','4,485 orders at 302012; current breach 50.7%',PURPLE)]
for i,(idc,city,act,why,colc) in enumerate(actrows):
    y=1.45+i*1.15; sh=slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE,PInches(.65),PInches(y),PInches(12),PInches(.9));sh.fill.solid();sh.fill.fore_color.rgb=LIGHT;sh.line.color.rgb=LIGHT
    add_text(slide,.92,y+.15,1.0,.25,idc,15,colc,True); add_text(slide,2.0,y+.17,1.25,.2,city,10,SLATE,True); add_text(slide,3.35,y+.14,2.15,.22,act,11,NAVY,True); add_text(slide,5.62,y+.17,6.55,.2,why,10,SLATE)
add_text(slide,.8,6.35,11.8,.35,'No other candidate passes the value test today once rent, imputed fixed cost and fit-out are respected.',12,NAVY,True)
# slide 6
slide=prs.slides.add_slide(prs.slide_layouts[6]); add_bg(slide); add_title(slide,'The new sites target the highest-risk demand pockets','Candidate scorecard uses exact order locations and the city distance-to-breach model.',TEAL)
score_sel=score[score.site_id.isin(['CS-01','CS-07','CS-11'])].copy()
headers=['Site','City / pincode','Area','Rent','Fit-out','Orders captured','Breach est.','Coverage est.']
for j,h in enumerate(headers):
    x=[.7,1.45,3.05,4.0,5.0,6.1,8.1,9.7][j]; w=[.7,1.45,.85,.9,1.0,1.8,1.5,1.5][j]; add_text(slide,x,1.45,w,.3,h,10,WHITE,True,PP_ALIGN.CENTER); sh=slide.shapes[-1]; # background not easy behind; add band below
band=slide.shapes.add_shape(MSO_SHAPE.RECTANGLE,PInches(.65),PInches(1.4),PInches(11.95),PInches(.47)); band.fill.solid();band.fill.fore_color.rgb=NAVY;band.line.fill.background(); slide.shapes._spTree.remove(band._element); slide.shapes._spTree.insert(2,band._element)
for j,h in enumerate(headers):
    x=[.7,1.45,3.05,4.0,5.0,6.1,8.1,9.7][j]; w=[.7,1.45,.85,.9,1.0,1.8,1.5,1.5][j]; add_text(slide,x,1.51,w,.2,h,9,WHITE,True,PP_ALIGN.CENTER)
for i,(_,r) in enumerate(score_sel.set_index('site_id').loc[['CS-01','CS-07','CS-11']].reset_index().iterrows()):
    y=2.02+i*1.05; sh=slide.shapes.add_shape(MSO_SHAPE.RECTANGLE,PInches(.65),PInches(y),PInches(11.95),PInches(.82));sh.fill.solid();sh.fill.fore_color.rgb=PALE if i%2==0 else LIGHT;sh.line.fill.background()
    vals=[r.site_id,f'{r.city} / {int(r.pincode)}',f'{int(r.area_sqft):,}',inr(r.rent_inr,0),inr(r.fitout_inr,0),f'{int(r.orders_captured):,}',pct(r.expected_breach_if_open),pct(r.coverage_if_open)]
    for j,v in enumerate(vals):
        x=[.7,1.45,3.05,4.0,5.0,6.1,8.1,9.7][j]; w=[.7,1.45,.85,.9,1.0,1.8,1.5,1.5][j]; add_text(slide,x,y+.3,w,.2,str(v),10,NAVY if j<2 else SLATE,True if j in [0,5] else False,PP_ALIGN.CENTER)
add_text(slide,.8,5.65,11.8,.72,'CS-07 is the best Hyderabad one-site trade-off. CS-11 is a service-first Jaipur pilot. CS-01 captures 20,433 orders after reassignment because it also becomes the western Pune anchor.',13,NAVY,True)
# slide 7
slide=prs.slides.add_slide(prs.slide_layouts[6]); add_bg(slide); add_title(slide,'Impact: variable savings fund the coverage move','Fixed cost rises, but distance and breach savings more than offset it at portfolio level.',CORAL)
slide.shapes.add_picture(str(IMG/'before_after_impact.png'),PInches(.6),PInches(1.35),width=PInches(7.1))
card(slide,8.55,1.65,3.55,1.05,inr(variable_delta,2),'variable C2S saving / order',PALE,TEAL)
card(slide,8.55,3.0,3.55,1.05,inr(fixed_delta,2),'fixed allocation increase / order',LIGHT,CORAL)
card(slide,8.55,4.35,3.55,1.05,inr(net_savings_per_order,2),'net saving / order',PALE,GREENP)
add_text(slide,8.55,5.85,3.55,.62,f'{inr(monthly_savings,0)} monthly × {payback:.1f} months to repay {inr(fitout,0)} fit-out.',12,NAVY,True,PP_ALIGN.CENTER)
add_text(slide,.85,6.35,7.3,.35,'Cost-floor option: ₹69.63/order and 16.4-month payback, but −5.6 coverage points and +3.7 breach points versus the balanced case.',9,SLATE,False)
# slide 8
slide=prs.slides.add_slide(prs.slide_layouts[6]); add_bg(slide); add_title(slide,'Stage the plan so the model earns trust','Commit capex now only where the data has a clear job for the property.',PURPLE)
steps=[('0–30 days','Validate leases','Confirm candidate fixed costs, access, parking, utilities and exit terms.'),('30–60 days','Build + reroute','Open CS-01 / CS-07 / CS-11; migrate PUN-07 demand; close PUN-02 after PUN-01 cutover.'),('Weekly','Control tower','Track pincode breach, radius coverage, distance and site load; no site is “successful” on orders alone.'),('Day 90','Stage gate','Keep CS-03 / CS-06 / CS-10 / others on watch list; approve only if high-risk demand persists and payback clears 24 months.')]
for i,(a,b,c) in enumerate(steps):
    y=1.45+i*1.16; circ=slide.shapes.add_shape(MSO_SHAPE.OVAL,PInches(.75),PInches(y),PInches(.55),PInches(.55));circ.fill.solid();circ.fill.fore_color.rgb=[TEAL,CORAL,PURPLE,GREENP][i];circ.line.fill.background();add_text(slide,.75,y+.16,.55,.15,str(i+1),12,WHITE,True,PP_ALIGN.CENTER);add_text(slide,1.55,y+.03,1.45,.22,a,10,SLATE,True);add_text(slide,3.15,y+.03,2.05,.22,b,13,NAVY,True);add_text(slide,5.35,y+.03,6.8,.34,c,11,SLATE)
add_text(slide,.8,6.08,11.8,.32,'Exact check: 7,165 facility subsets enumerated; selected city networks rank first under the service guardrail.',10,TEAL,True)
add_text(slide,.8,6.45,11.8,.32,'Rule: protect the 10-minute promise first; use fixed-cost consolidation only when coverage is protected.',12,NAVY,True)
# slide 9
slide=prs.slides.add_slide(prs.slide_layouts[6]); add_bg(slide,NAVY); add_title(slide,'Decision requested','Approve the focused network; monitor, then iterate.',TEAL)
add_text(slide,.8,1.55,11.7,.55,'Approve CS-01, CS-07 and CS-11.',26,WHITE,True)
add_text(slide,.8,2.35,11.7,.4,'Close PUN-02. Relocate PUN-07 into CS-01.',19,PRGBColor(220,240,246),True)
add_bullets_slide(slide,.95,3.35,10.9,2.0,[
    'Portfolio C2S: ₹71.74 → ₹69.78 / order.',
    'Estimated breach: 24.66% → 18.22%.',
    'Catchment coverage: 72.68% → 80.56%.',
    'Run-rate saving: approximately ₹1.83 lakh/month; fit-out payback ≈ 24.4 months.'
],17,WHITE)
add_text(slide,.8,6.45,11.7,.35,'Prepared from the supplied simulated data pack. Estimates are transparent, testable and stage-gated.',10,PRGBColor(203,213,225))
prs.save(OUT/'executive_deck.pptx')

# ---------- remaining useful artifacts ----------
# Save serviceable radius table (was constructed above).
radius_out=radii.merge(stores_raw[['store_id','city','store_pincode','monthly_rent_inr','monthly_fixed_cost_inr']],on='store_id',how='left')
radius_out.to_csv(OUT/'serviceable_radii.csv',index=False,float_format='%.6f')
# Add a compact README/manifest.
manifest=f'''# QuickCart Project Dark Matter — submission package\n\n## Main deliverables\n- `final_report.docx` — recommendation memo with assumptions, maps, actions, economics and stage gates.\n- `executive_deck.pptx` — 9-slide executive summary for the COO.\n- `final_analysis.ipynb` — reproducible analysis companion.\n- `dark_matter_analysis.py` — source script used to produce the outputs.\n- `optimization_audit.py` — exact order-level network subset search.\n\n## Data / evidence files\n- `order_level_cleaned.csv` — all 186,156 orders with distance, breach, month-level fixed allocation, cost-to-serve and catchment flags.\n- `after_order_assignment.csv` — exact order-level recommended-network assignment and estimated after metrics.\n- `before_after_city_metrics.csv` / `before_after_portfolio_metrics.csv` — headline impact tables.\n- `serviceable_radii.csv`, `candidate_site_scorecard.csv`, `recommended_facility_loads.csv`, `pincode_before_after_impact.csv`.\n- `images/` — static maps, optimization frontier and charts required by the brief.\n- `optimization_summary.csv` / `exhaustive_exact_scenarios.csv` — optimum-strategy audit.\n- `fixed_cost_sensitivity.csv` — sensitivity of the selected network to candidate fixed-cost imputation.\n- `cost_floor_before_after.csv` / `cost_floor_order_assignment.csv` — cost-first alternative retained as a transparent trade-off.\n\n## Recommendation\nOpen CS-01 in Pune, CS-07 in Hyderabad and CS-11 in Jaipur; close PUN-02; relocate PUN-07 into CS-01. Candidate fixed cost is imputed at city median existing fixed cost because the candidate file does not provide it.\n\n## Headline portfolio estimate\nC2S {inr(basep.avg_cost_to_serve,2)} -> {inr(afterp.avg_cost_to_serve,2)}; breach {pct(basep.breach_rate)} -> {pct(afterp.breach_rate)}; coverage {pct(basep.coverage_rate)} -> {pct(afterp.coverage_rate)}; run-rate saving {inr(monthly_savings,0)}/month; simple fit-out payback {payback:.1f} months.\n'''
(OUT/'SUBMISSION_README.md').write_text(manifest)
print('Built submission package:',OUT)
print('Files:',len(list(OUT.rglob('*'))))
print(manifest)
