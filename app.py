# -*- coding: utf-8 -*-
"""CompactoLab — Mesures de compactage, observations et incertitudes explicites."""
from __future__ import annotations

import csv
import hashlib
import io
import json
import math
import re
import unicodedata
import zipfile
from datetime import date, datetime, timezone
from html import escape
from pathlib import Path

import altair as alt
import numpy as np
import pandas as pd
from scipy import stats
import streamlit as st
from openpyxl import Workbook

APP_NAME='CompactoLab'
VERSION='1.0'
FOREST='#173F35';RUST='#A94326';SAND='#F4EBDD';INK='#233D34';GOLD='#987016'
COLS=['id','zone','date','couche','sol','reference','rho_d','rho_h','eau_pct','rho_ref','x_m','y_m','note']
NUMS=['rho_d','rho_h','eau_pct','rho_ref','x_m','y_m']
LABELS={'id':'Essai','zone':'Zone','date':'Date','couche':'Couche','sol':'Sol','reference':'Référence laboratoire','rho_d':'Masse volumique sèche (kg/m³)','rho_h':'Masse volumique humide (kg/m³)','eau_pct':'Teneur en eau (%)','rho_ref':'Masse volumique sèche de référence (kg/m³)','x_m':'X local (m)','y_m':'Y local (m)','note':'Observation','cr':'Compactage relatif (%)','methode':'Origine du calcul','ligne':'Ligne source'}
MAX_ROWS=10000;MAX_BYTES=5*1024*1024
FHWA_URL='https://www.fhwa.dot.gov/engineering/geotech/pubs/05037/05a.cfm'
NIST_URL='https://www.itl.nist.gov/div898/handbook/prc/section2/prc241.htm'


def fmt(value,digits=2):
    if value is None or pd.isna(value):return 'Non calculable'
    return f'{float(value):,.{digits}f}'.replace(',',' ').replace('.',',')


def name_key(value):
    s=unicodedata.normalize('NFKD',str(value)).encode('ascii','ignore').decode().lower()
    return re.sub('[^a-z0-9]','',s)


def standardize(frame):
    if not isinstance(frame,pd.DataFrame) or len(frame)>MAX_ROWS or len(frame.columns)>40:
        raise ValueError('Importer un tableau de 10 000 lignes et 40 colonnes maximum.')
    alias={name_key(k):k for k in COLS}
    for k,v in LABELS.items():
        if k in COLS:alias[name_key(v)]=k
    alias.update(essai='id',identifiant='id',densiteseche='rho_d',densitehumide='rho_h',teneureneau='eau_pct',humidite='eau_pct',densitereference='rho_ref',rhodref='rho_ref',x='x_m',y='y_m',commentaire='note')
    mapped={};ignored=[]
    for column in frame.columns:
        target=alias.get(name_key(column))
        if target:
            if target in mapped.values():raise ValueError('Deux colonnes représentent la même variable : '+target)
            mapped[column]=target
        else:ignored.append(str(column))
    if 'zone' not in mapped.values() or 'rho_ref' not in mapped.values() or not {'rho_d','rho_h'}&set(mapped.values()):
        raise ValueError('Colonnes requises : zone, rho_ref et rho_d ; ou zone, rho_ref, rho_h et eau_pct.')
    d=frame.rename(columns=mapped).reindex(columns=COLS).copy()
    for key in COLS:
        def text(v):
            if pd.isna(v):return ''
            if key=='date' and isinstance(v,(datetime,date,pd.Timestamp)):return v.strftime('%Y-%m-%d')
            return str(v).strip()
        d[key]=d[key].map(text)
        limit=1000 if key=='note' else 200
        if d[key].map(lambda v:len(v.encode('utf-8'))>limit).any():raise ValueError('Cellule trop longue dans '+key+f' (limite {limit} octets UTF-8).')
    if len(d.to_json(orient='records',force_ascii=False).encode('utf-8'))>10*1024*1024:
        raise ValueError('Carnet trop volumineux en mémoire : réduire les observations textuelles ou le nombre de lignes.')
    return d.reset_index(drop=True),ignored


def numeric(series):
    return pd.to_numeric(series.str.replace('\u00a0','',regex=False).str.replace(' ','',regex=False).str.replace(',','.',regex=False),errors='coerce').replace([np.inf,-np.inf],np.nan)


def clean_data(raw,dedup=False):
    if not len(raw):
        d=pd.DataFrame(columns=COLS+['ligne','cr','methode']);return d,pd.DataFrame(columns=['ligne','décision','motif']),dict(raw=0,valid=0,excluded=0,duplicates=0)
    d=raw.copy();errors=[[] for _ in range(len(d))]
    for key in NUMS:
        d[key]=numeric(raw[key]);supplied=raw[key].ne('')
        invalid=supplied&d[key].isna()
        if key in ['rho_d','rho_h','rho_ref']:invalid|=supplied&((d[key]<=0)|(d[key]>1e6))
        elif key=='eau_pct':invalid|=supplied&((d[key]<0)|(d[key]>1000))
        else:invalid|=supplied&d[key].abs().gt(1e7)
        for i in np.flatnonzero(invalid.to_numpy()):errors[i].append(key+' : valeur invalide')
    for i,row in raw.iterrows():
        if not row.zone:errors[i].append('Zone manquante')
        if not row.rho_ref:errors[i].append('Masse volumique de référence manquante')
        if row.rho_d and row.rho_h:errors[i].append('Saisir rho_d OU rho_h, pas les deux')
        if not row.rho_d and not (row.rho_h and row.eau_pct!=''):errors[i].append('Renseigner rho_d ou le couple rho_h + eau_pct')
        if row.date:
            try:
                if date.fromisoformat(row.date).isoformat()!=row.date:raise ValueError()
            except ValueError:errors[i].append('Date attendue au format AAAA-MM-JJ')
    humid=raw.rho_d.eq('')
    d.loc[humid,'rho_d']=d.loc[humid,'rho_h']/(1+d.loc[humid,'eau_pct']/100)
    d['methode']=np.where(humid,'Calculée : humide + eau','Sèche renseignée')
    d['cr']=100*d.rho_d/d.rho_ref
    bad_ratio=~np.isfinite(d.cr)|d.cr.le(0)|d.cr.gt(1e8)
    for i in np.flatnonzero(bad_ratio.to_numpy()):errors[i].append('Rapport de compactage non calculable')
    invalid=np.array([bool(e) for e in errors]);duplicate=raw.duplicated(keep='first').to_numpy()&~invalid
    if dedup:
        for i in np.flatnonzero(duplicate):errors[i].append('Doublon complet retiré sur demande')
    keep=np.array([not e for e in errors]);d['ligne']=np.arange(1,len(raw)+1)
    d['id']=[v if v else f'L-{i+1:04d}' for i,v in enumerate(d.id)]
    for key in ['couche','sol','reference']:d[key]=d[key].replace('','Non renseigné')
    audit=raw.copy();audit.insert(0,'ligne',np.arange(1,len(raw)+1));audit['décision']=['Exclue' if e else 'Conservée' for e in errors];audit['motif']=[' ; '.join(e) for e in errors]
    return d.loc[keep].reset_index(drop=True),audit,dict(raw=len(raw),valid=int(keep.sum()),excluded=int((~keep).sum()),duplicates=int(duplicate.sum()))


def below(values,target):
    # Neutralise uniquement les poussières d’arrondi du calcul, pas une tolérance de réception.
    return np.round(np.asarray(values,dtype=float),10)<round(float(target),10)


def describe(values):
    x=np.asarray(values,dtype=float);x=x[np.isfinite(x)];n=len(x)
    if not n:return dict(n=0,mean=None,median=None,std=None,variance=None,cv=None,min=None,max=None)
    sd=float(x.std(ddof=1)) if n>1 else None;mean=float(x.mean())
    return dict(n=n,mean=mean,median=float(np.median(x)),std=sd,variance=None if sd is None else sd**2,cv=100*sd/mean if sd is not None and mean>0 else None,min=float(x.min()),max=float(x.max()))


def probability(values,target):
    x=np.asarray(values,dtype=float);x=x[np.isfinite(x)]
    if not len(x):return None
    k=int(below(x,target).sum());n=len(x);s=describe(x)
    interval=stats.binomtest(k,n).proportion_ci(confidence_level=.95,method='wilson')
    normal=float(stats.norm.cdf(target,loc=s['mean'],scale=s['std'])) if n>=3 and s['std'] and s['std']>0 else None
    return dict(n=n,k=k,observed=k/n,low=float(interval.low),high=float(interval.high),normal=normal)


@st.cache_data(ttl=1800,max_entries=8,show_spinner=False)
def bootstrap_mean(values,repeats=2000,seed=42):
    x=np.asarray(values,dtype=float)
    if not 2<=len(x)<=MAX_ROWS or not np.isfinite(x).all() or repeats not in [1000,2000,5000] or not 0<=seed<=2147483647:
        raise ValueError('Bootstrap : au moins 2 valeurs finies, 1 000/2 000/5 000 tirages et graine autorisée.')
    rng=np.random.default_rng(seed);means=np.empty(repeats)
    for start in range(0,repeats,32):
        size=min(32,repeats-start);means[start:start+size]=x[rng.integers(0,len(x),size=(size,len(x)))].mean(axis=1)
    lo,hi=np.quantile(means,[.025,.975]);return dict(low=float(lo),high=float(hi),means=means,seed=seed,repeats=repeats)


def zone_table(df,target):
    rows=[]
    for zone,g in df.groupby('zone',sort=True):
        s=describe(g.cr);k=int(below(g.cr,target).sum())
        rows.append({'Zone':zone,'Essais':len(g),'Moyenne (%)':s['mean'],'Médiane (%)':s['median'],'Écart-type (points)':s['std'],'Sous objectif':k,'Part sous objectif (%)':100*k/len(g),'Références distinctes':len(g[['sol','reference','rho_ref','couche']].drop_duplicates())})
    return pd.DataFrame(rows,columns=['Zone','Essais','Moyenne (%)','Médiane (%)','Écart-type (points)','Sous objectif','Part sous objectif (%)','Références distinctes'])


def demo_data():
    cr=np.array([92,94,95,95,96,96,97,98,99,100],dtype=float)
    d=pd.DataFrame(dict(id=[f'F-{i+1:02d}' for i in range(10)],zone=['Zone A']*5+['Zone B']*5,date=['2026-10-15']*10,couche=['Couche d’exercice']*10,sol=['Sol fictif A']*10,reference=['REF-FICTIVE-01']*10,rho_d=cr*19,rho_ref=[1900]*10,eau_pct=[8,9,10,11,12,9,10,11,12,13],x_m=[0,10,20,30,40,50,60,70,80,90],y_m=[0,10]*5,note=['Exemple entièrement fictif']*10))
    return standardize(d)[0]


def read_upload(content,name):
    if not content or len(content)>MAX_BYTES:raise ValueError('Fichier vide ou supérieur à 5 Mo.')
    if name.lower().endswith('.xlsx'):
        try:
            with zipfile.ZipFile(io.BytesIO(content)) as z:
                if len(z.infolist())>500 or sum(x.file_size for x in z.infolist())>30*1024*1024:raise ValueError('Classeur décompressé trop volumineux.')
            frame=pd.read_excel(io.BytesIO(content),sheet_name=0,dtype=object,keep_default_na=False,engine='openpyxl',nrows=MAX_ROWS+1)
        except ValueError:raise
        except Exception as e:raise ValueError('Classeur XLSX illisible. Utiliser un fichier Excel récent non protégé.') from e
    elif name.lower().endswith('.csv'):
        try:text=content.decode('utf-8-sig')
        except UnicodeDecodeError:text=content.decode('cp1252')
        first=text.splitlines()[0] if text.splitlines() else ''
        sep=';' if first.count(';')>first.count(',') else '\t' if first.count('\t')>first.count(',') else ','
        try:frame=pd.read_csv(io.StringIO(text),sep=sep,dtype=str,keep_default_na=False,nrows=MAX_ROWS+1,on_bad_lines='error')
        except Exception as e:raise ValueError('CSV illisible : vérifier séparateur, en-têtes et nombre de cellules par ligne.') from e
    else:raise ValueError('Utiliser un CSV ou un XLSX ; convertir les anciens fichiers XLS.')
    return standardize(frame)


def safe_csv(df):
    d=df.copy()
    for col in d.select_dtypes(include=['object','string']).columns:
        d[col]=d[col].map(lambda v:"'"+v if isinstance(v,str) and v.lstrip().startswith(('=','+','-','@')) else v)
    return d.to_csv(index=False,sep=';',decimal=',',float_format='%.10g').encode('utf-8-sig')


def backup_bytes(ctx):
    payload=dict(application=APP_NAME,version=1,raw=ctx['raw'].to_dict('records'),meta=ctx['meta'],settings=ctx['settings'])
    return json.dumps(payload,ensure_ascii=False,allow_nan=False).encode('utf-8')


def selection_export(df):
    """Sélection réimportable : valeur sèche résolue, mesure humide tracée séparément."""
    out=df.copy();out['rho_h_mesuree']=out['rho_h'];out['rho_h']=''
    return out


def read_backup(content):
    if len(content)>12*1024*1024:raise ValueError('Sauvegarde limitée à 12 Mo.')
    try:
        o=json.loads(content)
        if not isinstance(o,dict) or o.get('application')!=APP_NAME or o.get('version')!=1:raise ValueError()
        rows=o['raw'];meta=o['meta'];settings=o['settings']
        if not isinstance(rows,list) or len(rows)>MAX_ROWS or not isinstance(meta,dict) or not isinstance(settings,dict):raise ValueError()
        if not all(isinstance(r,dict) and set(r)==set(COLS) and all(isinstance(v,str) for v in r.values()) for r in rows):raise ValueError()
        raw,_=standardize(pd.DataFrame(rows,columns=COLS))
        if set(meta)!=set(['name','kind','created']) or any(not isinstance(v,str) or len(v)>250 for v in meta.values()):raise ValueError()
        if set(settings)!=set(['target','zones','layers','dedup']):raise ValueError()
        if type(settings['dedup']) is not bool or isinstance(settings['target'],bool) or not isinstance(settings['target'],(int,float)) or not math.isfinite(settings['target']) or not 30<=settings['target']<=150:raise ValueError()
        for k in ['zones','layers']:
            if not isinstance(settings[k],list) or len(settings[k])>MAX_ROWS or not all(isinstance(v,str) and len(v)<=200 for v in settings[k]):raise ValueError()
        if meta['kind']=='Exemple fictif' and not raw.equals(demo_data()):meta['kind']='Origine déclarée'
        return raw,meta,settings
    except (ValueError,KeyError,TypeError,UnicodeError) as e:raise ValueError('Sauvegarde CompactoLab invalide ou incompatible.') from e


def metadata(name,kind):return dict(name=name,kind=kind,created=datetime.now(timezone.utc).isoformat(timespec='seconds'))


def activate(raw,meta,settings=None):
    st.session_state.cl_raw=raw.copy();st.session_state.cl_meta=meta;st.session_state.cl_revision=st.session_state.get('cl_revision',0)+1
    st.session_state.cl_defaults=settings or dict(target=95.,zones=[],layers=[],dedup=False)
    for key in ['cl_boot','cl_report','cl_risk']:st.session_state.pop(key,None)


def load_demo():activate(demo_data(),metadata('Démonstration · dix essais fictifs','Exemple fictif'))


def current():return st.session_state.cl_context


def signature(ctx):
    return hashlib.sha256(json.dumps(dict(raw=ctx['raw'].to_dict('records'),meta=ctx['meta'],settings=ctx['settings']),ensure_ascii=False,sort_keys=True).encode()).hexdigest()


def require_data():
    ctx=current()
    if ctx['df'].empty:
        st.info('Aucun essai analysable dans cette sélection. Charger des essais, vérifier le journal de nettoyage ou élargir les filtres.');return None
    return ctx


def heading(step,title,body):st.html(f'<div class="cl-heading"><span>{escape(step)}</span><h1>{escape(title)}</h1><p>{escape(body)}</p></div>')


def note(title,body,rust=False):
    st.html('<aside class="cl-note '+('rust' if rust else '')+'"><strong>'+escape(title)+'</strong><p>'+escape(body)+'</p></aside>')


def metric_row(items):
    st.html('<div class="cl-kpis">'+''.join('<article><span>'+escape(a)+'</span><b>'+escape(b)+'</b><small>'+escape(c)+'</small></article>' for a,b,c in items)+'</div>')


def table(df,alt,height='auto'):st.dataframe(df,hide_index=True,width='stretch',height=height,alt=alt)


def chart(c,alt,height=300):
    c=c.properties(height=height).configure_axis(labelColor=INK,titleColor=INK,gridColor='#E4D9C7',labelFontSize=11).configure_view(strokeWidth=0).configure_legend(labelColor=INK,titleColor=INK)
    st.altair_chart(c,width='stretch',theme=None,alt=alt)


def target_line(target):return alt.Chart(pd.DataFrame({'Objectif':[target]})).mark_rule(color=RUST,strokeDash=[6,3],strokeWidth=2).encode(y='Objectif:Q',tooltip='Objectif:Q')


def soil_section():
    st.html('<div class="cl-section" role="img" aria-label="Schéma pédagogique de trois couches de sol, sans échelle"><div class="cl-grass">Surface du terrain <span>MESURES PONCTUELLES</span></div><div class="cl-layer one"><i>01</i><strong>Couche étudiée</strong><span>Eau, masse volumique, référence</span></div><div class="cl-layer two"><i>02</i><strong>Couche inférieure</strong><span>À caractériser séparément</span></div><div class="cl-layer three"><i>03</i><strong>Sol support</strong><span>Schéma illustratif, sans épaisseur réelle</span></div></div>')


def set_zone(zone):st.session_state['cl_zones_'+str(st.session_state.cl_revision)]=[zone] if zone else []


def dashboard_page():
    ctx=current();d=ctx['df'];target=ctx['settings']['target']
    st.html('<header class="cl-hero"><div><span class="cl-eyebrow">CARNET D’OBSERVATION · GÉNIE CIVIL</span><h1>Le terrain,<br>mesure par mesure.</h1><p>Comparer les points contrôlés et comprendre leur variabilité.</p></div><div class="cl-stamp"><span>OBJECTIF D’EXERCICE</span><b>'+fmt(target,0)+'<small> %</small></b><p>Choisi par l’utilisateur<br>aucune norme imposée</p></div></header>')
    if ctx['raw'].empty:
        with st.container(horizontal=True,gap='medium'):
            st.button('Charger les 10 essais fictifs',on_click=load_demo,type='primary',icon=':material/science:',width='stretch',wrap=True)
            st.page_link(st.session_state.cl_pages['essais'],label='Saisir ou importer mes essais',icon=':material/edit_note:',width='stretch')
        a,b=st.columns([1.1,1])
        with a:
            st.subheader('Un projet que vous pouvez vérifier')
            note('Observer','Saisir la masse volumique sèche et sa référence, ou calculer la masse volumique sèche à partir d’une mesure humide et de la teneur en eau.')
            note('Comparer','Regarder les résultats par zone et par couche ; distinguer une mesure isolée d’une tendance.')
            note('Expliquer','Avec dix essais fictifs, deux sont sous l’objectif de 95 % : fréquence observée de 20 %.')
        with b:soil_section()
        st.caption('L’exemple est activé uniquement par votre bouton. Les données restent en session ; une sauvegarde JSON permet de les conserver.')
        return
    if d.empty:require_data();return
    s=describe(d.cr);p=probability(d.cr,target)
    metric_row([('Points sélectionnés',fmt(len(d),0),str(d.zone.nunique())+' zone(s) représentée(s)'),('Compactage moyen',fmt(s['mean'])+' %','Moyenne des rapports individuels'),('Sous l’objectif',f"{p['k']} / {p['n']}",fmt(100*p['observed'],1)+' % des essais observés'),('Dispersion',fmt(s['std'])+' pts' if s['std'] is not None else 'Non calculable','Écart-type d’échantillon (n−1)')])
    a,b=st.columns([1.35,1])
    with a:
        st.subheader('Les points de votre carnet')
        points=d.dropna(subset=['x_m','y_m']).copy();points['Lecture']=np.where(below(points.cr,target),'Sous l’objectif','Objectif atteint ou dépassé')
        if len(points):
            xmin,xmax=points.x_m.min(),points.x_m.max();ymin,ymax=points.y_m.min(),points.y_m.max();span=max(xmax-xmin,ymax-ymin,10)*1.12;cx=(xmin+xmax)/2;cy=(ymin+ymax)/2
            chart(alt.Chart(points).mark_circle(size=160,stroke='#FFFFFF',strokeWidth=1.5).encode(x=alt.X('x_m:Q',title='X local (m)',scale=alt.Scale(domain=[cx-span/2,cx+span/2],zero=False)),y=alt.Y('y_m:Q',title='Y local (m)',scale=alt.Scale(domain=[cy-span/2,cy+span/2],zero=False)),color=alt.Color('Lecture:N',scale=alt.Scale(domain=['Sous l’objectif','Objectif atteint ou dépassé'],range=[RUST,FOREST]),legend=alt.Legend(orient='bottom')),tooltip=['id:N','zone:N',alt.Tooltip('cr:Q',title='Compactage (%)',format='.2f'),'x_m:Q','y_m:Q']).interactive(),'Plan local des mesures avec coordonnées renseignées',320)
            st.caption(f'{len(points)} point(s) localisé(s) sur {len(d)}. Repère local saisi, sans carte GPS ni interpolation ; les axes ont la même étendue numérique, mais le dessin s’adapte à l’écran.')
        else:note('Pas encore de coordonnées','Renseignez X et Y, en mètres, pour situer les mesures. Le plan représente des points observés et ne certifie pas l’état entre ces points.')
    with b:
        st.subheader('Lire la couche étudiée');soil_section()
    st.subheader('Accéder aux zones')
    z=zone_table(ctx['base'],target);st.caption('Ces cartes représentent les zones des couches sélectionnées, avant le filtre de zone. Cliquer sur une zone actualise toutes les analyses.')
    zones=z.to_dict('records');page=0
    if len(zones)>9:page=st.number_input('Page des zones',1,math.ceil(len(zones)/9),1)-1
    for begin in range(0,len(zones[page*9:page*9+9]),3):
        cols=st.columns(3)
        for col,row in zip(cols,zones[page*9+begin:page*9+begin+3]):
            with col:
                with st.container(border=True):
                    st.markdown('**'+escape(row['Zone'])+'**')
                    st.metric('Moyenne observée',fmt(row['Moyenne (%)'])+' %')
                    st.caption(f"{row['Essais']} essais · {row['Sous objectif']} sous l’objectif")
                    st.button('Voir cette zone',key='cl_zonecard_'+row['Zone'],on_click=set_zone,args=(row['Zone'],),icon=':material/location_on:',width='stretch')
    st.button('Afficher toutes les zones',on_click=set_zone,args=(None,),icon=':material/select_all:')
    if (d.cr>100).any():note('Au-delà de 100 %',str(int(d.cr.gt(100).sum()))+' résultat(s) dépassent la référence déclarée. Ils sont conservés : vérifier référence, protocole et mesures avant d’interpréter.',True)
    note('Une observation, pas une couverture complète','Le pourcentage sous l’objectif concerne les essais sélectionnés. Il ne mesure ni une proportion de surface défectueuse ni une probabilité de défaillance de l’ouvrage.')


def add_record_page(ctx):
    st.subheader('Ajouter un essai')
    mode=st.segmented_control('Valeur disponible',['Masse volumique sèche','Masse volumique humide + eau'],default='Masse volumique sèche',key='cl_entry_mode')
    with st.form('cl_add_form',clear_on_submit=False):
        a,b=st.columns(2)
        ident=a.text_input('Identifiant de l’essai',max_chars=60,placeholder='E-001');zone=b.text_input('Zone *',max_chars=80,placeholder='Zone A')
        dt=a.date_input('Date de l’essai');layer=b.text_input('Couche',max_chars=80,placeholder='Couche 1')
        soil=a.text_input('Identification du sol',max_chars=80);reference=b.text_input('Référence de l’essai laboratoire',max_chars=80,help='Identifie une référence et son protocole. Une référence par type de sol et condition de laboratoire.')
        rho=a.number_input(('Masse volumique sèche' if mode=='Masse volumique sèche' else 'Masse volumique humide')+' (kg/m³)',min_value=1.,max_value=1e6,value=1800.,step=10.)
        ref=b.number_input('Masse volumique sèche de référence (kg/m³)',min_value=1.,max_value=1e6,value=1900.,step=10.)
        wet_known=st.checkbox('Teneur en eau connue',value=mode!='Masse volumique sèche',disabled=mode!='Masse volumique sèche')
        water=st.number_input('Teneur en eau (%)',min_value=0.,max_value=1000.,value=10.,step=.5,disabled=not wet_known and mode=='Masse volumique sèche')
        coords=st.checkbox('Ajouter des coordonnées locales')
        x=a.number_input('X local (m)',min_value=-1e7,max_value=1e7,value=0.,step=1.);y=b.number_input('Y local (m)',min_value=-1e7,max_value=1e7,value=0.,step=1.)
        comment=st.text_area('Observation (facultatif)',max_chars=400)
        submitted=st.form_submit_button('Enregistrer cet essai',type='primary')
    if submitted:
        if not zone.strip():st.error('Renseigner une zone avant d’enregistrer.');return
        row=dict(id=ident,zone=zone,date=dt.isoformat(),couche=layer,sol=soil,reference=reference,rho_ref=ref,eau_pct=water if (wet_known or mode!='Masse volumique sèche') else '',x_m=x if coords else '',y_m=y if coords else '',note=comment)
        row['rho_d' if mode=='Masse volumique sèche' else 'rho_h']=rho
        try:
            new,_=standardize(pd.DataFrame([row]));clean,audit,q=clean_data(new)
            if not len(clean):st.error(' ; '.join(audit.motif));return
            if len(ctx['raw'])>=MAX_ROWS:st.error('Limite de 10 000 essais atteinte.');return
            settings=dict(ctx['settings'],zones=[],layers=[])
            activate(pd.concat([ctx['raw'],new],ignore_index=True),metadata('Carnet saisi ou modifié','Données utilisateur'),settings);st.session_state.cl_flash='Essai ajouté. Les filtres de zones et couches ont été élargis pour le rendre visible.';st.rerun()
        except ValueError as e:st.error(str(e))


def data_page():
    heading('01 / CARNET','Vos mesures, leur origine et leurs unités.','Deux possibilités : renseigner directement la masse volumique sèche, ou renseigner la masse volumique humide et la teneur en eau.')
    ctx=current();rev=st.session_state.cl_revision
    with st.container(horizontal=True):
        st.button('Remplacer par les 10 essais fictifs',on_click=load_demo,icon=':material/science:',wrap=True)
        st.download_button('Télécharger le modèle CSV',safe_csv(demo_data().iloc[:0]),'compactolab-modele.csv','text/csv',on_click='ignore',icon=':material/download:',wrap=True)
    with st.expander('Importer un CSV ou un classeur XLSX'):
        upload=st.file_uploader('Fichier d’essais (5 Mo maximum)',type=['csv','xlsx'],max_upload_size=5,key='cl_import')
        if upload:
            try:
                raw,ignored=read_upload(upload.getvalue(),upload.name);clean,audit,q=clean_data(raw)
                st.write(f"{q['raw']} lignes ; {q['valid']} analysables ; {q['excluded']} exclues selon les règles.")
                if ignored:st.caption('Colonnes ignorées : '+', '.join(ignored))
                table(raw.head(10),'Aperçu du fichier importé')
                confirm=st.checkbox('Remplacer le carnet actif par ce fichier',key='cl_import_confirm')
                if st.button('Utiliser ce fichier',disabled=not confirm,type='primary'):
                    activate(raw,metadata(upload.name,'Import utilisateur'));st.rerun()
            except ValueError as e:st.error(str(e))
    add_record_page(ctx)
    with st.expander('Calcul à la main : de la masse au compactage relatif'):
        mass=st.number_input('Masse humide de sol seul (kg)',min_value=.0001,max_value=1e6,value=1.8,step=.1,key='cl_calc_mass')
        volume=st.number_input('Volume mesuré (litres)',min_value=.0001,max_value=1e6,value=1.,step=.1,key='cl_calc_volume')
        water=st.number_input('Teneur en eau pour le calcul (%)',0.,1000.,10.,key='cl_calc_water');ref=st.number_input('Référence pour le calcul (kg/m³)',1.,1e6,1800.,key='cl_calc_ref')
        rh=mass/(volume/1000);rd=rh/(1+water/100);cr=100*rd/ref
        st.latex(r'\rho_h=\frac{m_h}{V},\quad\rho_d=\frac{\rho_h}{1+w/100},\quad C_R=100\frac{\rho_d}{\rho_{d,ref}}')
        st.write(f"Volume = {fmt(volume/1000,6)} m³ ; masse volumique humide = {fmt(rh)} kg/m³ ; sèche = {fmt(rd)} kg/m³ ; compactage relatif = {fmt(cr)} %.")
        st.caption('La masse doit être celle du sol seul, après retrait de la tare. Eau (%) = masse d’eau / masse sèche × 100. Ce calcul ne s’ajoute pas automatiquement au carnet.')
    if len(ctx['raw']):
        st.subheader('Contrôle du carnet')
        q=ctx['quality'];metric_row([('Lignes brutes',str(q['raw']),'Avant nettoyage'),('Analysables',str(q['valid']),'Avant les filtres'),('Exclues',str(q['excluded']),'Motifs dans le journal'),('Doublons complets',str(q['duplicates']),'Conservés par défaut')])
        table(ctx['df'].rename(columns=LABELS),'Essais de la sélection active',height=330)
        with st.expander('Modifier le tableau brut ou supprimer des lignes'):
            if len(ctx['raw'])>1500:st.info('Au-delà de 1 500 lignes, exporter le CSV brut, le modifier et le réimporter.')
            else:
                st.caption('rho_d et rho_h s’excluent. Date : AAAA-MM-JJ. Vider une cellule facultative ne la remplace pas par zéro.')
                with st.form('cl_edit_'+str(rev)):
                    edited=st.data_editor(ctx['raw'],num_rows='dynamic',hide_index=True,width='stretch',alt='Tableau brut modifiable',key='cl_editor_'+str(rev))
                    confirm=st.checkbox('Appliquer ce tableau au carnet actif',key='cl_editor_ok_'+str(rev))
                    save=st.form_submit_button('Enregistrer le tableau')
                if save:
                    if not confirm:st.info('Cocher la confirmation pour appliquer le tableau.')
                    else:
                        try:
                            edited=edited.replace('',np.nan).dropna(how='all').fillna('');new,_=standardize(edited)
                            activate(new,metadata('Tableau utilisateur corrigé','Données utilisateur'),dict(ctx['settings'],zones=[],layers=[]));st.rerun()
                        except ValueError as e:st.error(str(e))
        with st.expander('Journal de nettoyage et export des valeurs brutes'):
            table(ctx['audit'],'Journal des lignes conservées et exclues',height=350)
            st.download_button('Valeurs brutes CSV',safe_csv(ctx['raw']),'compactolab-brut.csv','text/csv',on_click='ignore')
            st.download_button('Journal CSV',safe_csv(ctx['audit']),'compactolab-nettoyage.csv','text/csv',on_click='ignore')


def compare_page():
    heading('02 / ZONES','Comparer les zones sans masquer la dispersion.','Les résultats restent descriptifs. Les sols, les références de laboratoire et les couches doivent être comparables pour interpréter un écart.')
    ctx=require_data()
    if ctx is None:return
    d=ctx['df'];target=ctx['settings']['target'];z=zone_table(d,target)
    table(z.round(4),'Synthèse du compactage par zone')
    a,b=st.columns(2)
    with a:
        st.subheader('Tous les essais, par zone')
        chart(alt.Chart(d).mark_boxplot(color=FOREST,extent=1.5,size=36).encode(x=alt.X('zone:N',title=None,axis=alt.Axis(labelAngle=-25,labelLimit=150)),y=alt.Y('cr:Q',title='Compactage relatif (%)',scale=alt.Scale(zero=False)))+target_line(target),'Dispersion du compactage par zone et objectif choisi')
    with b:
        st.subheader('L’eau et la masse volumique sèche')
        wet=d.dropna(subset=['eau_pct'])
        if len(wet):
            chart(alt.Chart(wet).mark_circle(color=RUST,size=85,opacity=.8).encode(x=alt.X('eau_pct:Q',title='Teneur en eau (%)',scale=alt.Scale(zero=False)),y=alt.Y('rho_d:Q',title='Masse volumique sèche (kg/m³)',scale=alt.Scale(zero=False)),tooltip=['id:N','zone:N','eau_pct:Q','rho_d:Q']),'Mesures de terrain : teneur en eau et masse volumique sèche')
            st.caption('Ce nuage de mesures de terrain n’est pas une courbe Proctor de laboratoire. Aucun optimum en eau n’est déduit automatiquement.')
        else:note('Eau non renseignée','La comparaison des compactages reste disponible ; ajoutez des teneurs en eau pour explorer ce nuage de points.')
    zones=sorted(d.zone.unique())
    if len(zones)<2:st.info('Sélectionner au moins deux zones dans la barre latérale pour ouvrir le comparateur A / B.');return
    st.subheader('Face à face')
    a,b=st.columns(2);za=a.selectbox('Zone A',zones,key='cl_compare_a');zb=b.selectbox('Zone B',[v for v in zones if v!=za],key='cl_compare_b')
    da,db=d.loc[d.zone.eq(za)],d.loc[d.zone.eq(zb)];sa,sb=describe(da.cr),describe(db.cr)
    metric_row([(za,fmt(sa['mean'])+' %',str(sa['n'])+' essais'),(zb,fmt(sb['mean'])+' %',str(sb['n'])+' essais'),('Écart B − A',fmt(sb['mean']-sa['mean'])+' pts','Différence de moyennes'),('Écart-type A / B',fmt(sa['std'])+' / '+fmt(sb['std']),'En points de pourcentage')])
    note('Interprétation',f"Dans cette sélection, la moyenne de {zb} est {'supérieure' if sb['mean']>=sa['mean'] else 'inférieure'} à celle de {za} de {fmt(abs(sb['mean']-sa['mean']))} points. Les tailles des groupes et leurs références peuvent différer. Cet écart ne prouve pas la cause des différences.")
    st.download_button('Synthèse des zones CSV',safe_csv(z),'compactolab-zones.csv','text/csv',on_click='ignore',icon=':material/download:')


def cohorts(df):
    result=[]
    eligible=df.loc[df.sol.ne('Non renseigné')&df.reference.ne('Non renseigné')&df.couche.ne('Non renseigné')]
    for (soil,reference,rho,layer),g in eligible.groupby(['sol','reference','rho_ref','couche'],sort=True):
        label=f'{soil} · {reference} · {fmt(rho,0)} kg/m³ · {layer}'
        result.append((label,g.copy()))
    return result


def risk_page():
    heading('03 / INCERTITUDES','Ce que disent les essais, ce qu’ils ne disent pas.','La fréquence observée est un comptage. Une estimation probabiliste ajoute des hypothèses sur les essais et leur représentativité.')
    ctx=require_data()
    if ctx is None:return
    d=ctx['df'];target=ctx['settings']['target'];p=probability(d.cr,target)
    metric_row([('Sous l’objectif',str(p['k'])+' / '+str(p['n']),'Condition stricte CR < objectif'),('Fréquence observée',fmt(100*p['observed'],1)+' %','Sur les points sélectionnés'),('Objectif choisi',fmt(target)+' %','Repère d’exercice modifiable'),('Nombre de zones',str(d.zone.nunique()),'Pas une mesure de surface')])
    counts,edges=np.histogram(d.cr,bins=min(16,max(3,int(np.sqrt(len(d))))))
    hist=alt.Chart(d).mark_bar(color=FOREST,cornerRadiusTopLeft=3).encode(x=alt.X('cr:Q',bin=alt.Bin(maxbins=16),title='Compactage relatif (%)'),y=alt.Y('count():Q',title='Essais'))
    line=alt.Chart(pd.DataFrame({'objectif':[target]})).mark_rule(color=RUST,strokeWidth=2,strokeDash=[6,3]).encode(x='objectif:Q')
    chart(hist+line,'Distribution observée et objectif exploratoire')
    st.subheader('Estimer sous des hypothèses explicites')
    groups=cohorts(d)
    if not groups:
        st.session_state.pop('cl_risk',None)
        note('Identifier la référence avant de modéliser','Renseigner sol, référence laboratoire et couche dans le carnet. Les estimations seront limitées aux essais ayant ces informations communes et la même masse volumique de référence.',True);return
    labels=[v[0] for v in groups]
    label=st.selectbox('Sous-ensemble de même sol, couche et référence',labels,index=0 if len(labels)==1 else None,placeholder='Choisir un sous-ensemble',key='cl_cohort')
    if label is None:st.session_state.pop('cl_risk',None);st.info('Choisir un sous-ensemble pour éviter de mélanger des références différentes.');return
    group=groups[labels.index(label)][1]
    st.caption(f'{len(group)} essais dans ce sous-ensemble, sur {len(d)} de la sélection. La référence est considérée fixe et exacte pour ces calculs.')
    hypothesis=st.checkbox('Étudier le modèle en supposant des essais comparables, indépendants et représentatifs',key='cl_hypothesis',help='Hypothèse de travail, pas une vérification automatique. Des points proches ou des répétitions d’un même lot peuvent être dépendants.')
    if not hypothesis:st.session_state.pop('cl_risk',None);return
    result=probability(group.cr,target);s=describe(group.cr)
    st.session_state.cl_risk=dict(signature=signature(ctx),label=label,probability=result,summary=s)
    metric_row([('Fréquence du sous-ensemble',fmt(100*result['observed'],1)+' %',f"{result['k']} / {result['n']} essais"),('Intervalle de Wilson à 95 %',fmt(100*result['low'],1)+' – '+fmt(100*result['high'],1)+' %','Incertitude sur une proportion'),('Modèle normal ajusté',fmt(100*result['normal'],1)+' %' if result['normal'] is not None else 'Non ajustable','Probabilité théorique sous l’objectif'),('Dispersion du sous-ensemble',fmt(s['std'])+' pts' if s['std'] is not None else 'Non calculable','Écart-type avec diviseur n−1')])
    note('Bien lire l’intervalle','L’intervalle de Wilson décrit une procédure d’estimation d’une proportion sous les hypothèses annoncées. Il ne signifie pas que 95 % des compactages sont dans cette plage et ne mesure pas une probabilité de défaillance de l’ouvrage.')
    if result['normal'] is not None:
        lo=min(float(group.cr.min()),s['mean']-3*s['std'],target);hi=max(float(group.cr.max()),s['mean']+3*s['std'],target);xs=np.linspace(lo,hi,240)
        curve=pd.DataFrame({'Compactage (%)':xs,'Densité':stats.norm.pdf(xs,loc=s['mean'],scale=s['std'])})
        density=alt.Chart(curve).mark_line(color=FOREST,strokeWidth=2).encode(x='Compactage (%):Q',y='Densité:Q')
        shade=alt.Chart(curve.loc[curve['Compactage (%)']<target]).mark_area(color=RUST,opacity=.32).encode(x='Compactage (%):Q',y='Densité:Q')
        chart(shade+density+line,'Modèle normal : surface colorée sous l’objectif choisi',240)
        st.caption('La normale est une hypothèse à discuter, pas une loi prouvée. Elle est non bornée et peut attribuer une masse à des compactages physiquement impossibles. Un échantillon d’au moins trois valeurs et une dispersion non nulle sont exigés pour l’afficher.')
    else:st.info('Le modèle normal exige au moins trois valeurs et une dispersion non nulle. Une série constante ne démontre pas une absence de risque.')
    with st.expander('Incertitude sur la moyenne : bootstrap'):
        st.caption('Rééchantillonnage des compactages observés avec remise. L’intervalle concerne la moyenne du sous-ensemble, pas les valeurs individuelles. La référence reste fixe.')
        with st.form('cl_boot_form'):
            n=st.select_slider('Rééchantillonnages',[1000,2000,5000],value=2000);seed=st.number_input('Graine aléatoire',0,2147483647,42)
            run=st.form_submit_button('Estimer la moyenne par bootstrap',disabled=len(group)<2,type='primary')
        sig=signature(ctx)+'|'+label
        if run:
            with st.spinner('Rééchantillonnages…'):st.session_state.cl_boot=(sig,bootstrap_mean(tuple(group.cr),n,seed))
        boot=st.session_state.get('cl_boot')
        if boot and boot[0]==sig:
            r=boot[1];note('Intervalle percentile à 95 %',f"Moyenne entre {fmt(r['low'])} et {fmt(r['high'])} %. {r['repeats']} rééchantillonnages ; graine {r['seed']}.")
            chart(alt.Chart(pd.DataFrame({'Moyenne (%)':r['means']})).mark_bar(color=FOREST).encode(x=alt.X('Moyenne (%):Q',bin=alt.Bin(maxbins=22),title='Moyenne du compactage (%)'),y=alt.Y('count():Q',title='Rééchantillonnages')),'Moyennes obtenues par rééchantillonnage',240)


def xlsx_bytes(ctx):
    wb=Workbook();wb.remove(wb.active)
    context=pd.DataFrame([{'Champ':k,'Valeur':str(v)} for k,v in {**ctx['meta'],**ctx['settings']}.items()])
    for name,frame in [('Selection',selection_export(ctx['df'])),('Zones',zone_table(ctx['df'],ctx['settings']['target'])),('Nettoyage',ctx['audit']),('Brut',ctx['raw']),('Contexte',context)]:
        ws=wb.create_sheet(name);ws.append([str(c) for c in frame.columns])
        for row in frame.itertuples(index=False,name=None):
            ws.append([None if pd.isna(v) else v.item() if isinstance(v,np.generic) else v for v in row])
            for cell in ws[ws.max_row]:
                if isinstance(cell.value,str):cell.data_type='s'
        ws.freeze_panes='A2';ws.auto_filter.ref=ws.dimensions
        for col in ws.columns:ws.column_dimensions[col[0].column_letter].width=min(44,max(14,len(str(col[0].value))+3))
    out=io.BytesIO();wb.save(out);return out.getvalue()


def report_conclusion(ctx):
    d=ctx['df'];p=probability(d.cr,ctx['settings']['target']);s=describe(d.cr)
    if not p:return 'Aucun essai analysable dans la sélection ; vérifier les données, les exclusions et les filtres.'
    return f"La sélection comprend {p['n']} essais sur {d.zone.nunique()} zone(s). Le compactage relatif moyen est de {fmt(s['mean'])} %. {p['k']} essais, soit {fmt(100*p['observed'])} %, sont strictement sous l’objectif exploratoire de {fmt(ctx['settings']['target'])} %. Cette fréquence décrit les points observés ; elle ne mesure pas l’étendue des zones concernées. Les références et l’indépendance des essais doivent être examinées avant toute généralisation."


def pdf_figure(df,target,by_zone=False):
    from reportlab.graphics.shapes import Drawing,Rect,Line,String
    from reportlab.lib import colors
    from reportlab.pdfbase.pdfmetrics import stringWidth
    draw=Drawing(515,235);draw.add(Rect(0,0,515,235,fillColor=colors.HexColor('#F5EEDD'),strokeColor=None,rx=10,ry=10))
    if df.empty:
        draw.add(String(30,115,'Aucun essai analysable.',fontName='CLFont',fontSize=11));return draw
    if by_zone:
        z=zone_table(df,target).head(12);vals=z['Moyenne (%)'].tolist();labels=z.Zone.tolist()
    else:sample=df.head(30);vals=sample.cr.tolist();labels=sample.id.tolist()
    maximum=max(max(vals),target)*1.12;left,right,bottom,top=44,495,52,208
    Y=lambda v:bottom+float(v)/maximum*(top-bottom)
    for v in np.linspace(0,maximum,5):
        draw.add(Line(left,Y(v),right,Y(v),strokeColor=colors.HexColor('#DED0BA'),strokeWidth=.5));draw.add(String(left-7,Y(v)-3,fmt(v,1),fontName='CLFont',fontSize=7,textAnchor='end',fillColor=colors.HexColor(INK)))
    step=(right-left)/len(vals)
    for i,(value,label) in enumerate(zip(vals,labels)):
        x=left+i*step+step*.16
        draw.add(Rect(x,bottom,step*.68,Y(value)-bottom,fillColor=colors.HexColor(RUST if below([value],target)[0] else FOREST),strokeColor=None))
        if len(vals)<=12:
            text=str(label)
            while len(text)>1 and stringWidth(text,'CLFont',6.5)>step*.86:text=text[:-2]+'…'
            draw.add(String(x+step*.34,bottom-15,text,fontName='CLFont',fontSize=6.5,textAnchor='middle',fillColor=colors.HexColor(INK)))
    draw.add(Line(left,Y(target),right,Y(target),strokeColor=colors.HexColor(RUST),strokeDashArray=[4,3],strokeWidth=1.2))
    draw.add(String(left,220,'Compactage relatif (%) · ligne pointillée = objectif choisi',fontName='CLFont',fontSize=8,fillColor=colors.HexColor(INK)))
    draw.add(String(left,16,'Moyennes par zone (12 premières)' if by_zone else 'Essais dans l’ordre du carnet (30 premiers)',fontName='CLFont',fontSize=8,fillColor=colors.HexColor(INK)))
    return draw


def pdf_bytes(ctx,options=None,risk=None):
    import reportlab
    from reportlab.lib import colors
    from reportlab.lib.pagesizes import A4
    from reportlab.lib.styles import getSampleStyleSheet,ParagraphStyle
    from reportlab.pdfbase import pdfmetrics
    from reportlab.pdfbase.ttfonts import TTFont
    from reportlab.platypus import SimpleDocTemplate,Paragraph,Spacer,Table,TableStyle,PageBreak,KeepTogether
    font_dir=Path(reportlab.__file__).parent/'fonts'
    if 'CLFont' not in pdfmetrics.getRegisteredFontNames():
        pdfmetrics.registerFont(TTFont('CLFont',str(font_dir/'Vera.ttf')));pdfmetrics.registerFont(TTFont('CLFont-Bold',str(font_dir/'VeraBd.ttf')))
        pdfmetrics.registerFontFamily('CLFont',normal='CLFont',bold='CLFont-Bold',italic='CLFont',boldItalic='CLFont-Bold')
    styles=getSampleStyleSheet()
    for style in styles.byName.values():style.fontName='CLFont';style.fontSize=9;style.leading=13
    styles['Title'].fontSize=24;styles['Title'].leading=29;styles['Title'].textColor=colors.HexColor(FOREST)
    styles['Heading2'].fontName='CLFont-Bold';styles['Heading2'].fontSize=12;styles['Heading2'].leading=17;styles['Heading2'].textColor=colors.HexColor(RUST)
    small=ParagraphStyle('CLSmall',fontName='CLFont',fontSize=8,leading=11)
    options=options or dict(title='Analyse pédagogique du compactage',student='',conclusion=report_conclusion(ctx))
    items=[];out=io.BytesIO();d=ctx['df'];target=ctx['settings']['target'];s=describe(d.cr);p=probability(d.cr,target)
    def para(text,style='BodyText'):items.extend([Paragraph(escape(str(text)),styles[style]),Spacer(1,7)])
    def tb(headers,rows,widths):
        t=Table([[Paragraph(escape(str(v)),small) for v in row] for row in [headers]+rows],colWidths=widths,repeatRows=1,hAlign='LEFT')
        t.setStyle(TableStyle([('BACKGROUND',(0,0),(-1,0),colors.HexColor('#E1E9DD')),('ROWBACKGROUNDS',(0,1),(-1,-1),[colors.white,colors.HexColor('#F6F0E6')]),('VALIGN',(0,0),(-1,-1),'TOP'),('TOPPADDING',(0,0),(-1,-1),7),('BOTTOMPADDING',(0,0),(-1,-1),7)]));items.extend([t,Spacer(1,12)])
    para('CompactoLab','Title');para(options['title'],'Heading2');para('Étudiant / groupe : '+(options.get('student') or 'Non renseigné'))
    para('Généré le '+datetime.now(timezone.utc).strftime('%Y-%m-%d à %H:%M UTC')+' · Version '+VERSION)
    para('Périmètre et provenance','Heading2');para(ctx['meta']['name']+' · '+ctx['meta']['kind'])
    para(f"Lignes brutes : {ctx['quality']['raw']} ; analysables avant filtres : {ctx['quality']['valid']} ; exclues : {ctx['quality']['excluded']} ; sélection : {len(d)}.")
    para('Zones : '+(', '.join(ctx['settings']['zones']) or 'toutes')+' ; couches : '+(', '.join(ctx['settings']['layers']) or 'toutes')+'. Doublons retirés : '+('oui' if ctx['settings']['dedup'] else 'non')+'.')
    para('Objectif exploratoire : '+fmt(target)+' %. Choix de l’utilisateur, sans validation réglementaire.')
    para('Indicateurs observés','Heading2')
    tb(['Effectif','Moyenne (%)','Médiane (%)','Écart-type (pts)','CV (%)'],[[str(s['n']),fmt(s['mean']),fmt(s['median']),fmt(s['std']),fmt(s['cv'])]],[65,110,110,120,80])
    if p:para(f"Strictement sous l’objectif : {p['k']}/{p['n']}, soit {fmt(100*p['observed'])} %. L’égalité à l’objectif n’est pas comptée sous le seuil.")
    else:para('Sélection vide : pas de moyenne ni de proportion calculable.')
    para('Formules et unités','Heading2');para('Masse volumique humide = masse humide de sol seul / volume. Volume en litres / 1000 = volume en m³. Masse volumique sèche = masse volumique humide / (1 + eau % / 100). Compactage relatif = 100 × masse volumique sèche / masse volumique sèche de référence. Masses volumiques en kg/m³.')
    para('Variance et écart-type d’échantillon avec diviseur n-1 (au moins deux essais). La référence est déclarée par l’utilisateur et considérée fixe. Le compactage relatif ne doit pas être confondu avec l’indice de densité des sols granulaires.')
    items.append(PageBreak());para('Résultats par zone','Heading2')
    z=zone_table(d,target)
    if len(z):tb(['Zone','n','Moyenne (%)','Écart-type (pts)','Sous objectif'],[[r['Zone'],str(r['Essais']),fmt(r['Moyenne (%)']),fmt(r['Écart-type (points)']),str(r['Sous objectif'])] for r in z.head(30).to_dict('records')],[145,40,100,110,90])
    else:para('Aucune zone dans la sélection.')
    if len(z)>30:para('Tableau limité aux 30 premières zones. Le classeur Excel contient la synthèse complète.')
    para('Les moyennes peuvent regrouper des références différentes. La comparaison ne prouve ni causalité ni uniformité spatiale et ne couvre pas les endroits non mesurés.')
    items.append(KeepTogether([Paragraph('Lecture graphique des zones',styles['Heading2']),pdf_figure(d,target,True),Spacer(1,12)]))
    items.append(PageBreak());para('Observations et interprétation','Heading2');items.append(pdf_figure(d,target))
    para('Le graphique affiche au plus 30 essais dans l’ordre du carnet, sans sélection des meilleurs résultats. Les indicateurs portent sur toute la sélection. Tous les essais sont exportables en CSV et Excel.')
    if risk and risk.get('signature')==signature(ctx):
        r=risk['probability'];para('Estimation probabiliste retenue','Heading2');para('Sous-ensemble : '+risk['label'])
        para(f"{r['n']} essais ; fréquence {fmt(100*r['observed'])} % ; intervalle de Wilson à 95 % : {fmt(100*r['low'])} à {fmt(100*r['high'])} %. Modèle normal ajusté : "+(fmt(100*r['normal'])+' % sous l’objectif.' if r['normal'] is not None else 'non calculable.'))
        para('Hypothèses activées par l’utilisateur : essais comparables, indépendants et représentatifs ; référence fixe. Le modèle normal est non borné. Ces résultats ne mesurent pas une probabilité de défaillance de l’ouvrage. Le bootstrap interactif n’est pas inclus dans le rapport.')
    else:para('Aucune estimation probabiliste retenue pour ce contexte ; les fréquences ci-dessus sont uniquement descriptives.')
    para('Conclusion relue par l’utilisateur','Heading2')
    for line in options['conclusion'].splitlines():
        if line.strip():para(line)
    para('Sources des méthodes et limites','Heading2');para('FHWA, Geotechnical Aspects of Pavements, chapitre 5, équations 5.1 et 5.4 : '+FHWA_URL)
    para('NIST/SEMATECH, intervalle de confiance de Wilson pour une proportion : '+NIST_URL)
    para('Les sources documentent les méthodes ; elles ne fournissent pas les données de démonstration. Les dix essais de démonstration sont entièrement fictifs. Aucun seuil de réception, protocole d’essai complet ni dimensionnement n’est certifié par cette application.')
    def footer(canvas,doc):
        canvas.setFont('CLFont',8);canvas.setFillColor(colors.HexColor(FOREST));canvas.drawString(36,22,'CompactoLab · Analyse pédagogique');canvas.drawRightString(A4[0]-36,22,str(doc.page))
    SimpleDocTemplate(out,pagesize=A4,leftMargin=36,rightMargin=36,topMargin=36,bottomMargin=44,title=options['title'],author=options.get('student') or APP_NAME).build(items,onFirstPage=footer,onLaterPages=footer)
    return out.getvalue()


def methods_block():
    st.subheader('Comprendre et tester en quelques minutes')
    with st.expander('Test guidé : dix essais fictifs',expanded=True):
        st.markdown('1. Charger les dix essais fictifs et conserver l’objectif **95 %**.\n2. Retrouver **2 / 10 sous l’objectif**, soit **20 %** ; moyenne **96,2 %**.\n3. Dans Comparer les zones : **A = 94,4 %**, **B = 98 %** ; écart B − A = **3,6 points**.\n4. Dans Probabilités : choisir le sous-ensemble commun, activer les hypothèses ; intervalle de Wilson ≈ **5,67–50,98 %**.\n5. Préparer et télécharger PDF, Excel et JSON. Changer l’objectif à **96 %** : **4 / 10**, soit **40 %**, passent strictement sous le seuil. Aucun résultat égal à 96 % n’est compté sous ce seuil.')
        st.button('Charger ce test fictif',on_click=load_demo,type='primary',icon=':material/science:')
    with st.expander('Formules et exemple de conversion'):
        st.latex(r'\rho_d=\frac{\rho_h}{1+w/100},\quad C_R=100\frac{\rho_d}{\rho_{d,ref}}')
        st.markdown('Pour **1,8 kg** de sol humide dans **1 litre** : ρh = **1 800 kg/m³**. Avec **10 %** d’eau : ρd = 1 800 / 1,10 = **1 636,36 kg/m³**. Pour une référence de **1 800 kg/m³**, CR = **90,91 %**. La teneur en eau est rapportée à la masse de sol sec.')
        st.markdown('La moyenne est la somme divisée par n. La variance d’échantillon utilise n−1 ; l’écart-type en est la racine. Le CV vaut 100 × écart-type / moyenne. Une seule mesure ne permet pas d’estimer une variance.')
    with st.expander('Limites et origine des méthodes'):
        st.markdown('Le seuil est un choix d’exercice. La référence doit correspondre au sol et au protocole identifiés. Une valeur supérieure à 100 % reste visible et invite à vérifier les informations. Aucune interpolation entre les points ni prévision d’une stabilité d’ouvrage n’est réalisée.\n\nUne fréquence décrit les mesures sélectionnées ; son extrapolation suppose un échantillonnage adapté. Les points voisins et les répétitions peuvent être dépendants. Les groupes peuvent différer par sol, eau, couche et référence. Le nuage eau/masse volumique n’est pas une courbe Proctor et ne donne pas d’optimum automatique.')
        st.markdown(f'[FHWA : relations masse volumique / eau et compactage relatif]({FHWA_URL}) · [NIST : intervalle de Wilson]({NIST_URL})')
        st.caption('Les références décrivent des méthodes générales ; aucune valeur normative locale n’est importée. Les exemples de l’application sont construits et fictifs.')


def reports_page():
    heading('04 / RAPPORTS & MÉTHODES','Conserver des résultats que vous pouvez expliquer.','Le PDF présente l’analyse ; le JSON conserve le carnet et ses réglages pour une restauration ultérieure.')
    ctx=current()
    if len(ctx['raw']):
        with st.container(horizontal=True):
            st.download_button('Sauvegarder le carnet JSON',backup_bytes(ctx),'compactolab-carnet.json','application/json',on_click='ignore',icon=':material/save:',width='stretch',wrap=True)
            st.download_button('Sélection CSV',safe_csv(selection_export(ctx['df'])),'compactolab-selection.csv','text/csv',on_click='ignore',icon=':material/download:',width='stretch',wrap=True)
        title=st.text_input('Titre du rapport',value='Analyse pédagogique du compactage',max_chars=180,key='cl_report_title',persist_state='session')
        student=st.text_input('Étudiant / groupe (facultatif)',max_chars=120,key='cl_report_student',persist_state='session')
        sig=signature(ctx);ckey='cl_conclusion_'+sig[:12]
        conclusion=st.text_area('Conclusion modifiable',value=report_conclusion(ctx),max_chars=3000,height=165,key=ckey,persist_state='session')
        risk=st.session_state.get('cl_risk');risk=risk if risk and risk['signature']==sig else None
        include=st.checkbox('Inclure l’estimation probabiliste étudiée',disabled=risk is None,key='cl_include_risk_'+sig[:12],persist_state='session')
        if risk:st.caption('Sous-ensemble étudié : '+risk['label']+' ; hypothèses explicites activées dans Probabilités.')
        else:st.caption('Pour inclure un intervalle ou un modèle : ouvrir Probabilités et activer les hypothèses sur un sous-ensemble identifié.')
        options=dict(title=title.strip() or 'Analyse pédagogique du compactage',student=student,conclusion=conclusion)
        selected_risk=risk if include else None
        rsig=hashlib.sha256((sig+json.dumps(options,ensure_ascii=False,sort_keys=True)+json.dumps(selected_risk,ensure_ascii=False,sort_keys=True)).encode()).hexdigest()
        if st.button('Préparer le PDF et Excel',type='primary',icon=':material/picture_as_pdf:'):
            with st.spinner('Préparation des documents…'):st.session_state.cl_report=(rsig,pdf_bytes(ctx,options,selected_risk),xlsx_bytes(ctx))
        saved=st.session_state.get('cl_report')
        if saved and saved[0]==rsig:
            with st.container(horizontal=True):
                st.download_button('Télécharger le PDF',saved[1],'compactolab-rapport.pdf','application/pdf',on_click='ignore',width='stretch')
                st.download_button('Télécharger Excel',saved[2],'compactolab-analyse.xlsx','application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',on_click='ignore',width='stretch')
        elif saved:st.info('Les données, les filtres ou les textes ont changé : préparer les documents à nouveau.')
        st.caption('Excel : Selection, Zones, Nettoyage, Brut, Contexte. Les textes du rapport et le bootstrap ne sont pas enregistrés dans le JSON ; ils restent dans la session. Aucune base de données distante n’est utilisée.')
    with st.expander('Restaurer un carnet sauvegardé'):
        up=st.file_uploader('Sauvegarde CompactoLab JSON (12 Mo maximum)',type='json',max_upload_size=12,key='cl_restore')
        if up:
            try:
                raw,meta,settings=read_backup(up.getvalue());st.write(f"{len(raw)} lignes · {meta['name']} · {meta['kind']}")
                ok=st.checkbox('Remplacer le carnet et ses réglages par cette sauvegarde',key='cl_restore_confirm')
                if st.button('Restaurer le carnet',disabled=not ok):activate(raw,meta,settings);st.rerun()
            except ValueError as e:st.error(str(e))
    methods_block()


def apply_style():
    st.html('''<style>
    :root{--cl-forest:#173F35;--cl-rust:#A94326;--cl-sand:#F4EBDD;--cl-ink:#233D34}
    [data-testid="stAppViewContainer"]{background:radial-gradient(ellipse at 100% 0,#E9DFCB,transparent 45%),#F8F4EB;color:var(--cl-ink);font-family:'Segoe UI',Arial,sans-serif}
    [data-testid="stMainBlockContainer"]{max-width:1440px;padding:4.5rem 2.5rem 3rem}
    [data-testid="stHeader"]{background:#F8F4EBF2}
    [data-testid="stSidebar"]{background:#E8DECB;border-right:1px solid #CFBEA3;color:#233D34}
    [data-testid="stSidebar"] label,[data-testid="stSidebar"] h3,[data-testid="stSidebarNavLink"] span{color:#233D34!important}
    [data-testid="stSidebarNavLink"]{min-height:42px;border-radius:6px}
    [data-testid="stSidebarNavLink"][aria-current="page"]{background:#173F35;border-left:4px solid #D7B36A}
    [data-testid="stSidebarNavLink"][aria-current="page"] span{color:#FFF8E9!important}
    [data-testid="stCaptionContainer"]{opacity:1!important}[data-testid="stCaptionContainer"] p{color:#5D6254}
    [data-testid="stWidgetLabel"] p{white-space:normal!important;overflow:visible!important}
    [data-testid="stAppViewContainer"] h1,[data-testid="stAppViewContainer"] h2,[data-testid="stAppViewContainer"] h3{color:#233D34;font-family:'Segoe UI',Arial,sans-serif;letter-spacing:-.025em}
    [data-testid="stAppViewContainer"] button{font-family:'Segoe UI',Arial,sans-serif;font-weight:650}
    .stButton button,.stDownloadButton button{min-height:44px;border-radius:7px;border:1px solid #BFAE94}
    button[kind="primary"]{background:#A94326!important;color:#FFF9F2!important;border-color:#A94326!important}
    .stButton button:hover,.stDownloadButton button:hover{border-color:#A94326;box-shadow:0 4px 14px #7A45251C}
    [data-testid="stForm"],[data-testid="stExpander"]{border:1px solid #CFC1A9;background:#FFFCF5;border-radius:10px}
    [data-testid="stMetric"]{background:#FFFDF7;border-radius:9px;padding:12px 14px}
    .cl-brand{padding:20px 0 25px}.cl-brand>span{display:inline-flex;align-items:center;justify-content:center;width:46px;height:46px;background:#173F35;color:#F4DAA6;border-radius:8px;font-size:18px;font-weight:800;letter-spacing:-1px;margin-bottom:13px;box-shadow:5px 5px 0 #C89964}.cl-brand strong{display:block;font-size:27px;letter-spacing:-1.1px;color:#173F35}.cl-brand small{display:block;font-size:10px;letter-spacing:2px;color:#5B4B34;margin-top:5px}
    .cl-topline{display:flex;align-items:center;justify-content:space-between;gap:12px;padding-bottom:15px;border-bottom:1px solid #CFBEA3;margin-bottom:15px}.cl-topline>span{color:#655844;font-size:10px;letter-spacing:1.5px}.cl-chip{font-size:11px;padding:7px 12px;background:#E2E9D9;color:#234B35;border-radius:5px;max-width:100%;overflow-wrap:anywhere}
    .cl-hero{display:flex;align-items:center;justify-content:space-between;gap:25px;background:#173F35;border-radius:12px;padding:31px 34px;margin:10px 0 24px;color:#FFF9E8;border-right:7px solid #C87B4D;position:relative}
    .cl-eyebrow{color:#EAC88B;font-size:10px;letter-spacing:2px;font-weight:650}.cl-hero h1{color:#FFF9E8!important;font-family:Georgia,'Times New Roman',serif!important;font-size:47px;line-height:1.08;font-weight:500;letter-spacing:-1.4px;margin:12px 0 15px;padding:0}.cl-hero p{font-size:14px;line-height:1.7;color:#E3E9D6;margin:0}
    .cl-stamp{flex:0 0 175px;border:1px dashed #A9BD99;padding:18px 20px;text-align:center;border-radius:50% 50% 8px 8px;background:#214B3E}.cl-stamp>span{font-size:8px;letter-spacing:1.3px;color:#EAC88B}.cl-stamp b{display:block;font-family:Georgia,serif;font-size:55px;font-weight:500;color:#F0D29F;margin:6px 0}.cl-stamp b small{font-size:22px}.cl-stamp p{font-size:10px;line-height:1.6}
    .cl-kpis{display:grid;grid-template-columns:repeat(4,minmax(0,1fr));gap:13px;margin:18px 0 25px}.cl-kpis article{min-width:0;background:#FFFDF7;padding:17px 19px;border:1px solid #CFC1A9;border-radius:9px;border-bottom:4px solid #173F35}.cl-kpis article:nth-child(2){background:#E7EDDD;border-color:#BBCBAE}.cl-kpis article:nth-child(3){background:#F3E3D8;border-color:#C8977A;border-bottom-color:#A94326}.cl-kpis article:nth-child(4){border-bottom-color:#987016}.cl-kpis span{display:block;font-size:12px;color:#4E5B4D}.cl-kpis b{display:block;font-size:27px;line-height:1.3;margin:9px 0;color:#173F35;overflow-wrap:anywhere}.cl-kpis small{font-size:11px;color:#515A4E;line-height:1.55;display:block}
    .cl-heading>span{font-size:11px;letter-spacing:2px;color:#A94326;font-weight:650}.cl-heading h1{font-size:34px;line-height:1.18;margin:12px 0;padding:0}.cl-heading p{font-size:15px;line-height:1.65;color:#5D6254;margin:0 0 24px}
    .cl-note{border:1px solid #C2CDB8;border-left:4px solid #416349;background:#EBF0E2;border-radius:5px;padding:15px 19px;margin:16px 0}.cl-note strong{color:#234633;font-size:14px}.cl-note p{font-size:13px;color:#3C5444;line-height:1.65;margin:6px 0 0}.cl-note.rust{background:#F7E9DF;border-color:#D8B397;border-left-color:#A94326}.cl-note.rust strong,.cl-note.rust p{color:#7B3E2B}
    .cl-section{margin:10px 0 17px;border:1px solid #CDB795;overflow:hidden;border-radius:12px;background:#E3D4B6}.cl-grass{padding:15px 18px;color:#EAEEDB;background:#416349;font-size:12px;display:flex;justify-content:space-between;gap:8px}.cl-grass span{color:#E8C98E;font-size:8px;letter-spacing:1px}.cl-layer{padding:20px 18px;position:relative;border-bottom:1px solid #9A79534D;background-size:14px 14px}.cl-layer.one{background-color:#C98D61;background-image:radial-gradient(#8F553A40 1px,transparent 1px)}.cl-layer.two{background-color:#BDA879;background-image:radial-gradient(#6D623E45 1px,transparent 1px)}.cl-layer.three{background-color:#8F9278;background-image:radial-gradient(#46543B45 1px,transparent 1px)}.cl-layer i{font-style:normal;font-family:Georgia,serif;font-size:24px;opacity:.8;margin-right:15px;color:#2C382A;float:left}.cl-layer strong{display:block;font-size:14px;color:#263A2A}.cl-layer>span{display:block;font-size:11px;margin-top:5px;color:#283B2B}
    @media(max-width:1000px){.cl-kpis{grid-template-columns:repeat(2,minmax(0,1fr))}.cl-hero h1{font-size:40px}}
    @media(max-width:640px){[data-testid="stMainBlockContainer"]{padding:4.5rem 1rem 2rem}.cl-topline{align-items:flex-start;flex-direction:column;gap:9px}.cl-hero{padding:23px 20px;align-items:flex-start;flex-direction:column;gap:18px;border-right-width:5px}.cl-hero h1{font-size:38px}.cl-stamp{width:100%;flex-basis:auto;border-radius:7px;text-align:left;box-sizing:border-box;padding:11px 15px;display:grid;grid-template-columns:1fr 1fr;align-items:center;gap:3px 15px}.cl-stamp>span{grid-column:1}.cl-stamp b{grid-column:2;grid-row:1/3;font-size:40px;text-align:right}.cl-stamp p{grid-column:1;grid-row:2;font-size:9px}.cl-heading h1{font-size:28px}.cl-kpis article{padding:13px 11px}.cl-kpis b{font-size:23px}.cl-grass{flex-direction:column;gap:5px}.cl-layer{padding:17px 15px}}
    @media(max-width:350px){.cl-kpis{grid-template-columns:1fr}.cl-hero h1{font-size:33px}}
    </style>''')


def main():
    st.set_page_config(page_title=APP_NAME+' · Carnet des sols',page_icon=':material/landscape:',layout='wide',initial_sidebar_state='auto')
    for key,value in [('cl_raw',pd.DataFrame(columns=COLS)),('cl_meta',metadata('Aucun carnet','Aucune source')),('cl_revision',0),('cl_defaults',dict(target=95.,zones=[],layers=[],dedup=False))]:
        if key not in st.session_state:st.session_state[key]=value
    apply_style()
    pages=[st.Page(dashboard_page,title='Tableau de bord',icon=':material/space_dashboard:',default=True,url_path='accueil'),st.Page(data_page,title='Essais et saisie',icon=':material/edit_note:',url_path='essais'),st.Page(compare_page,title='Comparer les zones',icon=':material/compare_arrows:',url_path='zones'),st.Page(risk_page,title='Probabilités',icon=':material/query_stats:',url_path='probabilites'),st.Page(reports_page,title='Rapports et méthodes',icon=':material/description:',url_path='rapports')]
    st.session_state.cl_pages={p.url_path:p for p in pages}
    with st.sidebar:st.html('<div class="cl-brand"><span>CL</span><strong>CompactoLab</strong><small>LE CARNET DES SOLS</small></div>')
    page=st.navigation(pages,position='sidebar',expanded=True)
    raw=st.session_state.cl_raw;defaults=st.session_state.cl_defaults;rev=str(st.session_state.cl_revision)
    # Le nettoyage précède le filtrage : le journal reste complet et les exclusions visibles.
    clean,audit,q=clean_data(raw,False)
    with st.sidebar:
        st.subheader('Périmètre d’analyse')
        target=st.number_input('Objectif exploratoire (%)',min_value=30.,max_value=150.,value=float(defaults['target']),step=.5,key='cl_target_'+rev,help='95 % est une valeur initiale pour l’exercice, pas un seuil réglementaire.')
        layers=st.multiselect('Couches',sorted(clean.couche.unique()) if len(clean) else [],default=[v for v in defaults['layers'] if len(clean) and v in clean.couche.values],key='cl_layers_'+rev,placeholder='Toutes les couches')
        zones=st.multiselect('Zones',sorted(clean.zone.unique()) if len(clean) else [],default=[v for v in defaults['zones'] if len(clean) and v in clean.zone.values],key='cl_zones_'+rev,placeholder='Toutes les zones')
        dedup=st.checkbox('Retirer les doublons complets',value=defaults['dedup'],key='cl_dedup_'+rev,wrap=True)
        st.caption('Une sélection vide signifie toutes les zones ou toutes les couches. Les analyses concernent uniquement les essais, pas la surface du terrain.')
        st.caption('Version '+VERSION+' · Projet pédagogique')
    if dedup:clean,audit,q=clean_data(raw,True)
    base=clean.loc[clean.couche.isin(layers)].copy() if layers else clean.copy()
    d=base.loc[base.zone.isin(zones)].copy() if zones else base.copy()
    settings=dict(target=target,zones=zones,layers=layers,dedup=dedup)
    ctx=dict(raw=raw,clean=clean,base=base,df=d,audit=audit,quality=q,meta=st.session_state.cl_meta,settings=settings)
    st.session_state.cl_context=ctx
    st.html('<div class="cl-topline"><span>COMPACTOLAB / TERRAIN & MESURES</span><div class="cl-chip">'+escape(ctx['meta']['kind'])+' · '+str(len(d))+' essai(s)</div></div>')
    if 'cl_flash' in st.session_state:st.success(st.session_state.pop('cl_flash'))
    if len(raw):st.caption(ctx['meta']['name']+' · Les filtres s’appliquent aux cinq rubriques.')
    page.run()


if __name__=='__main__':main()
