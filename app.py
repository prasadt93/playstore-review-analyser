# ═══════════════════════════════════════════════════════════════════════════════
#  Play Store Review Analyser  ·  Streamlit App  v2
#  No API keys · No AI · Runs fully offline on your Mac
# ═══════════════════════════════════════════════════════════════════════════════

import streamlit as st
import json, os, re, io, warnings
from datetime import datetime, timedelta
from collections import Counter
import pandas as pd
import numpy as np

warnings.filterwarnings("ignore")

from google_play_scraper import reviews as gps_reviews, app as gps_app, Sort
from vaderSentiment.vaderSentiment import SentimentIntensityAnalyzer
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter
import plotly.express as px
import plotly.graph_objects as go

# ══════════════════════════════════════════════════════════════════════════════
#  ISSUE & FEATURE CATEGORY ENGINE
#  Uses universal problem-type templates that apply to ANY app.
#  Each category is only included if it has meaningful presence in the data.
#  This gives proper labels ("Login & OTP Problems") not random bigrams.
# ══════════════════════════════════════════════════════════════════════════════

# These templates work for any app — e-commerce, banking, gaming, fintech, etc.
# We only include a category if reviews actually mention it (data-driven filtering).
_UNIVERSAL_ISSUE_SEEDS = {
    "App Crashes & Force Close": [
        "crash","crashing","crashed","crashes","force close","force stop",
        "keeps stopping","stopped working","not opening","won't open",
        "black screen","blank screen","keeps crashing","auto close",
    ],
    "Slow Performance & Lag": [
        "slow","lag","lagging","laggy","sluggish","hangs","hanging",
        "freezing","frozen","takes forever","very slow","too slow",
        "not smooth","response time","buffering","takes time",
    ],
    "Login & Authentication Problems": [
        "login","log in","otp","sign in","password","unable to login",
        "cannot login","login failed","session expired","logged out",
        "keeps logging out","verification code","2fa","authentication",
        "login issue","cant login","sign in issue",
    ],
    "Payment & Transaction Failures": [
        "payment","transaction","payment failed","transaction failed",
        "payment not","payment error","unable to pay","payment issue",
        "payment declined","payment not processed","checkout",
    ],
    "Refund & Money Not Received": [
        "refund","money not received","amount not","refund not",
        "not refunded","money stuck","amount deducted","wrongly charged",
        "extra charge","charged","deducted","money not credited",
    ],
    "Fund Transfer & Withdrawal Issues": [
        "transfer","withdrawal","withdraw","bank transfer","funds not",
        "transfer failed","payout","imps","neft","rtgs","upi transfer",
        "money not transferred","transfer not",
    ],
    "Customer Support & Response": [
        "customer care","customer service","support","helpline","complaint",
        "no response","not responding","not helpful","agent","grievance",
        "no help","service is bad","unresolved","no callback","chat support",
    ],
    "UI & Navigation Issues": [
        "confusing","complicated","difficult to use","not user friendly",
        "navigation","bad design","hard to find","cluttered","layout",
        "interface","ui issue","not intuitive","poor design","ux",
    ],
    "Ads & Intrusive Pop-ups": [
        "ads","advertisement","popup","pop-up","pop up","too many ads",
        "intrusive","spam","promotional","banner","ad",
    ],
    "Account Access & KYC": [
        "account","account blocked","suspended","kyc","not verified",
        "account issue","profile","demat","cannot access","account not",
    ],
    "Data & Content Not Loading": [
        "not loading","not showing","data not","not refreshing",
        "wrong data","incorrect data","not updating","blank page",
        "content not","data missing","not displaying",
    ],
    "Server & Network Errors": [
        "server","server error","network","connection","internet",
        "server down","timeout","502","503","error code","no internet",
        "connection failed","offline",
    ],
    "App Update & Compatibility Issues": [
        "after update","new update","latest update","update broke",
        "since update","update issue","new version","version issue",
        "compatibility","downgrade",
    ],
    "Notification & Alert Problems": [
        "notification","push notification","not getting notification",
        "notification not","missed notification","alert not",
        "no notification","notification issue",
    ],
    "Order & Booking Issues": [
        "order","booking","order failed","order not","order issue",
        "order cancelled","booking failed","cant order","order error",
        "order not placed","order status",
    ],
    "Search & Filter Not Working": [
        "search","search not","cant find","not finding","filter",
        "search results","wrong results","search issue","no results",
    ],
    "Sync & Data Loss Issues": [
        "sync","syncing","data lost","data wiped","not syncing",
        "lost data","reset","erased","missing data","backup",
    ],
    "Feature Broken / Not Working": [
        "not working","feature not","broken","doesn't work","does not work",
        "stopped working","used to work","no longer","feature missing",
        "feature issue","bug",
    ],
}

_UNIVERSAL_FEATURE_SEEDS = {
    "Dark Mode / Night Theme": [
        "dark mode","night mode","dark theme","dark ui","black theme","amoled",
    ],
    "Biometric / Fingerprint Login": [
        "fingerprint","biometric","face id","face unlock","touch id","mpin",
    ],
    "Offline Mode": [
        "offline","offline mode","no internet mode","work offline","without internet",
    ],
    "Multiple Account Support": [
        "multiple account","multi account","switch account","add account","family account",
    ],
    "Price / Item Alerts": [
        "price alert","set alert","notify me","back in stock","price drop","stock alert",
    ],
    "Better Search & Filters": [
        "better search","advanced filter","search filter","more filter","search option",
    ],
    "Download / Export Data": [
        "download","export","pdf download","export data","download history",
        "download statement","download report",
    ],
    "Widget / Home Screen Shortcut": [
        "widget","home screen widget","shortcut","quick access","home widget",
    ],
    "Language Support": [
        "hindi","language","regional language","tamil","telugu","kannada",
        "add language","more language",
    ],
    "Better Customer Support": [
        "better support","live chat","human agent","call support","faster support",
        "improve support","24/7 support",
    ],
    "Tablet / iPad Layout": [
        "tablet","ipad","tablet view","landscape mode","tablet layout",
    ],
    "Two-Factor Authentication": [
        "two factor","2fa","extra security","secure login","pin lock","app lock",
    ],
    "Undo / History Feature": [
        "undo","redo","history","action history","undo option",
    ],
    "Sorting & Customisation": [
        "sort","custom","customise","customize","rearrange","reorder","personalize",
    ],
    "Faster Loading / Performance": [
        "faster","speed up","improve speed","better performance","reduce loading",
        "optimise","optimize",
    ],
}


def discover_issue_categories(df, min_pct=0.01):
    """
    Filter universal issue templates to only those with ≥ min_pct presence
    in the app's negative reviews. Returns them sorted by prevalence.
    Gives meaningful labels for any app, not random word fragments.
    """
    neg = df[df["score"] <= 2]["content"].dropna()
    if len(neg) < 5:
        return {}

    scored = {}
    for label, kws in _UNIVERSAL_ISSUE_SEEDS.items():
        hit = neg.apply(lambda x: contains_any(x, kws)).sum()
        pct = hit / len(neg)
        if pct >= min_pct:
            scored[label] = (kws, pct)

    # Sort by prevalence, return top 20
    return {label: kws for label, (kws, _) in
            sorted(scored.items(), key=lambda x: -x[1][1])[:20]}


def discover_feature_categories(df, min_mentions=2):
    """
    Filter universal feature templates to only those actually requested
    in the app's reviews. Returns them sorted by mention count.
    """
    all_text = df["content"].dropna()
    scored = {}
    for label, kws in _UNIVERSAL_FEATURE_SEEDS.items():
        count = int(all_text.apply(lambda x: contains_any(x, kws)).sum())
        if count >= min_mentions:
            scored[label] = (kws, count)

    return {label: kws for label, (kws, _) in
            sorted(scored.items(), key=lambda x: -x[1][1])}

# ══════════════════════════════════════════════════════════════════════════════
#  CONFIG
# ══════════════════════════════════════════════════════════════════════════════

HISTORY_FILE = os.path.join(os.path.dirname(os.path.abspath(__file__)), "review_history.json")
MONTHS_OPTIONS = {"Last 1 Month":1,"Last 3 Months":3,"Last 6 Months":6,"Last 12 Months":12}
STOPWORDS = {
    "the","a","an","and","or","but","in","on","at","to","for","of","with",
    "by","from","is","are","was","were","be","been","have","has","had","do",
    "does","did","will","would","could","should","may","might","can","this",
    "that","these","those","i","me","my","we","our","you","your","it","its",
    "not","no","so","if","as","up","out","just","very","also","using","use",
    "used","get","got","need","want","like","really","even","more","than",
    "when","now","all","one","new","time","after","before","they","their",
    "there","them","he","she","his","her","us","who","what","which","where",
    "how","why","any","some","make","much","many","most","other","into",
    "through","same","too","only","own","then","such","both","each","few",
    "over","once","while","still","well","back","give","see","look","take",
    "come","thing","things","think","feel","always","never","every","since",
    "already","yet","quite","seem","said","says","try","keep","going","making",
    "doing","working","works","work","update","updated","version","play",
    "google","store","review","star","stars","rating","phone","mobile",
    "device","app","please","good","great","nice","bad","worst","best",
}
sia = SentimentIntensityAnalyzer()

# ══════════════════════════════════════════════════════════════════════════════
#  HISTORY
# ══════════════════════════════════════════════════════════════════════════════

def load_history():
    if os.path.exists(HISTORY_FILE):
        try:
            with open(HISTORY_FILE,"r") as f:
                return json.load(f)
        except Exception:
            return []
    return []

def save_history(history):
    with open(HISTORY_FILE,"w") as f:
        json.dump(history,f,indent=2)

def add_to_history(app_id,app_name,url,icon_url=""):
    history=load_history()
    history=[h for h in history if h.get("app_id")!=app_id]
    history.insert(0,{"app_id":app_id,"app_name":app_name,"url":url,
                       "icon_url":icon_url,"last_searched":datetime.now().isoformat()})
    save_history(history[:25])

# ══════════════════════════════════════════════════════════════════════════════
#  SCRAPING
# ══════════════════════════════════════════════════════════════════════════════

def extract_app_id(url):
    m=re.search(r"[?&]id=([a-zA-Z0-9_.]+)",url)
    return m.group(1) if m else None

@st.cache_data(ttl=3600,show_spinner=False)
def fetch_app_info(app_id):
    for country in ["in","us"]:
        try:
            return gps_app(app_id,lang="en",country=country)
        except Exception:
            continue
    return {}

@st.cache_data(ttl=3600,show_spinner=False)
def fetch_reviews_cached(app_id,months):
    cutoff=datetime.now()-timedelta(days=months*30.5)
    all_reviews,token=[],None
    for _ in range(60):
        try:
            batch,token=gps_reviews(app_id,lang="en",country="in",
                                     sort=Sort.NEWEST,count=200,continuation_token=token)
        except Exception:
            break
        if not batch:
            break
        for r in batch:
            rev_date=r.get("at")
            if isinstance(rev_date,datetime) and rev_date<cutoff:
                return all_reviews
            all_reviews.append(r)
        if not token:
            break
    return all_reviews

# ══════════════════════════════════════════════════════════════════════════════
#  ANALYSIS ENGINE
# ══════════════════════════════════════════════════════════════════════════════

def contains_any(text, keywords):
    t = str(text).lower()
    return any(re.search(r'\b' + re.escape(kw) + r'\b', t) for kw in keywords)

def sent_label(text):
    s=sia.polarity_scores(str(text))["compound"]
    return "Positive" if s>=0.05 else ("Negative" if s<=-0.05 else "Neutral")

def classify_trend(vals):
    valid=[v for v in vals if v and v>0]
    if not valid or max(valid)<0.005:
        return "✅ Stable-Low"
    if max(valid)<0.02:
        return "✅ Stable-Low"
    n=len(vals)
    half=max(n//2,1)
    def smean(lst):
        c=[v for v in lst if v]
        return float(np.mean(c)) if c else 0.0
    first_avg=smean(vals[:half])
    last_avg =smean(vals[half:])
    avg_all  =smean(valid)
    if last_avg<first_avg*0.55 and first_avg>0.02:
        return "✅ Improved"
    if last_avg>first_avg*1.45 and last_avg>0.02:
        return "🔺 Rising"
    high_m=sum(1 for v in valid if v>avg_all*1.5 and v>0.05)
    if high_m>=2:
        return "⚡ Recurring Spike"
    if avg_all>=0.03:
        return "⚠ Persists"
    return "✅ Stable-Low"

def peak_month(vals,labels):
    if not vals or max(vals)<0.005:
        return "—"
    idx=int(np.argmax(vals))
    return f"{labels[idx]} ({vals[idx]*100:.0f}%)"

def analyse_issue_trends(df, issue_categories):
    """issue_categories is discovered dynamically from the data."""
    neg=df[df["score"]<=2].copy()
    neg["month"]=neg["at"].dt.to_period("M")
    months=sorted(neg["month"].unique())
    mlabels=[str(m) for m in months]
    rows={}
    for cat,kws in issue_categories.items():
        vals=[]
        for m in months:
            sub=neg[neg["month"]==m]
            if len(sub)==0:
                vals.append(0.0)
                continue
            hit=sub["content"].apply(lambda x:contains_any(x,kws)).sum()
            vals.append(round(hit/len(sub),4))
        rows[cat]=vals
    return pd.DataFrame(rows,index=mlabels).T if rows else pd.DataFrame()

def analyse_feature_requests(df, feature_categories):
    """feature_categories is discovered dynamically from the data."""
    df=df.copy()
    df["month"]=df["at"].dt.to_period("M")
    months=sorted(df["month"].unique())
    mlabels=[str(m) for m in months]
    rows={}
    for feat,kws in feature_categories.items():
        vals=[]
        for m in months:
            sub=df[df["month"]==m]
            if len(sub)==0:
                vals.append(None)
                continue
            hit=int(sub["content"].apply(lambda x:contains_any(x,kws)).sum())
            vals.append(hit if hit>0 else None)
        if any(v for v in vals if v):
            rows[feat]=vals
    return pd.DataFrame(rows,index=mlabels).T if rows else pd.DataFrame()

def generate_insights(issue_df,feat_df,monthly,app_name):
    mlabels=issue_df.columns.tolist()
    improved,persistent,spikes,rising=[],[],[],[]
    for cat in issue_df.index:
        vals=issue_df.loc[cat].tolist()
        valid=[v for v in vals if v and v>0]
        if not valid or max(valid)<0.01:
            continue
        trend=classify_trend(vals)
        n=len(vals)
        t3=max(n//3,1)
        def smean(lst):
            c=[v for v in lst if v]
            return float(np.mean(c)) if c else 0.0
        f3=smean(vals[:t3]); l3=smean(vals[-t3:])
        pm=peak_month(vals,mlabels)
        base={"issue":cat,"finding":"","detail":"","peak":pm}
        if trend=="✅ Improved":
            base["finding"]=f"Dropped from {f3*100:.0f}% → {l3*100:.0f}% of negative reviews"
            base["detail"] =f"Peaked at {pm}. Consistent decline in later months suggests the issue was addressed in a subsequent release."
            improved.append(base)
        elif trend=="⚠ Persists":
            avg=smean(valid)
            base["finding"]=f"Consistently {avg*100:.0f}–{max(valid)*100:.0f}% of negative reviews all period"
            base["detail"] ="No clear resolution trend. Present across every month. Requires structured engineering fix."
            persistent.append(base)
        elif trend=="⚡ Recurring Spike":
            base["finding"]=f"Peaked at {pm}; spiked multiple times"
            base["detail"] =f"Fluctuating pattern — likely triggered by app update releases. {f3*100:.0f}% early vs {l3*100:.0f}% recently."
            spikes.append(base)
        elif trend=="🔺 Rising":
            base["finding"]=f"Growing: {f3*100:.0f}% early → {l3*100:.0f}% recently"
            base["detail"] ="Issue was minor initially but has intensified. Needs immediate attention."
            rising.append(base)

    feat_insights=[]
    if feat_df is not None and not feat_df.empty:
        for feat in feat_df.index:
            vals=feat_df.loc[feat].fillna(0).tolist()
            total=int(sum(vals))
            mths=int(sum(1 for v in vals if v and v>0))
            if total>=2:
                feat_insights.append({"feature":feat,
                    "finding":f"Requested in {mths} month(s) — {total} total mentions","total":total})
        feat_insights.sort(key=lambda x:-x["total"])

    avg_rat=float(monthly["avg_rating"].mean()) if "avg_rating" in monthly.columns else 0
    strategic=[
        {"area":"Overall Health",
         "finding":f"{app_name} holds avg {avg_rat:.2f}★ — {len(improved)} issues resolved, {len(persistent)} ongoing",
         "detail":"While several issues improved, persistent complaints create a ceiling on ratings."},
        {"area":"Release Regression Risk",
         "finding":f"{len(spikes)} issue(s) show recurring spike patterns",
         "detail":"Simultaneous spikes in crashes, slowness, and core features suggest new releases lack adequate regression testing."},
        {"area":"Feature Competitiveness",
         "finding":f"{len(feat_insights)} distinct feature types requested by users",
         "detail":"Users are explicitly comparing the app to competitor apps. Closing feature gaps can meaningfully improve ratings."},
        {"area":"Quick Wins",
         "finding":"Specific, repeated feature asks are likely low-effort improvements",
         "detail":"Highly specific requests (e.g. chart indicators, alert features) suggest clear user expectations that can be met quickly."},
    ]
    return {"improved":improved,"persistent":persistent,"spikes":spikes,
            "rising":rising,"feature_requests":feat_insights,"strategic":strategic}

def analyse_reviews(df):
    df=df.copy()
    df["at"]=pd.to_datetime(df["at"])
    df["sentiment"]=df["content"].apply(sent_label)
    df["compound"] =df["content"].apply(lambda x:sia.polarity_scores(str(x))["compound"])

    results={"df":df,"total_reviews":len(df),
             "avg_rating":round(df["score"].mean(),2),
             "rating_dist":df["score"].value_counts().sort_index(),
             "sentiment_dist":df["sentiment"].value_counts()}

    replied=df["replyContent"].notna().sum() if "replyContent" in df.columns else 0
    results["reply_rate"]=round(replied/max(len(df),1)*100,1)

    df["month"]=df["at"].dt.to_period("M")
    monthly=df.groupby("month").agg(
        count=("score","count"),avg_rating=("score","mean"),
        one_star=("score",lambda x:(x==1).sum()),
        two_star=("score",lambda x:(x==2).sum()),
        three_star=("score",lambda x:(x==3).sum()),
        four_star=("score",lambda x:(x==4).sum()),
        five_star=("score",lambda x:(x==5).sum()),
        positive=("sentiment",lambda x:(x=="Positive").sum()),
        negative=("sentiment",lambda x:(x=="Negative").sum()),
        neutral=("sentiment", lambda x:(x=="Neutral").sum()),
    ).reset_index()
    monthly["month_str"]=monthly["month"].astype(str)
    monthly["avg_rating"]=monthly["avg_rating"].round(2)
    monthly["neg_pct"]=((monthly["one_star"]+monthly["two_star"])/monthly["count"]*100).round(1)
    monthly["pos_pct"]=((monthly["four_star"]+monthly["five_star"])/monthly["count"]*100).round(1)
    results["monthly"]=monthly

    last30=df[df["at"]>=datetime.now()-timedelta(days=30)]
    results["recent_avg"]=round(last30["score"].mean(),2) if len(last30) else results["avg_rating"]

    # Discover categories from actual data — unique to each app
    issue_cats   = discover_issue_categories(df)
    feature_cats = discover_feature_categories(df)
    results["issue_cats"] = issue_cats
    results["issue_df"]   = analyse_issue_trends(df, issue_cats)
    results["feature_df"] = analyse_feature_requests(df, feature_cats)

    # Tag each review with matching issue categories (only where relevant)
    def tag_issues(text):
        matches = [label for label, kws in issue_cats.items()
                   if contains_any(text, kws)]
        return ", ".join(matches) if matches else ""

    df["issue_tags"] = df["content"].apply(tag_issues)

    def extract_kw(texts,n=15):
        words=[]
        for t in texts:
            words.extend([w for w in re.findall(r"\b[a-z]{3,}\b",str(t).lower()) if w not in STOPWORDS])
        return Counter(words).most_common(n)

    results["positive_themes"]=extract_kw(df[df["sentiment"]=="Positive"]["content"])
    results["negative_themes"]=extract_kw(df[df["sentiment"]=="Negative"]["content"])
    results["best_reviews"] =df[df["sentiment"]=="Positive"].nlargest(5,"thumbsUpCount")[["content","score","thumbsUpCount","at"]]
    results["worst_reviews"]=df[df["sentiment"]=="Negative"].nsmallest(5,"compound")[["content","score","thumbsUpCount","at"]]
    return results

# ══════════════════════════════════════════════════════════════════════════════
#  EXCEL COLOUR HELPERS
# ══════════════════════════════════════════════════════════════════════════════

def issue_fill(val):
    if not val or val==0:     return PatternFill("solid",fgColor="FFFFFF")
    if val>=0.15:             return PatternFill("solid",fgColor="C00000")
    if val>=0.10:             return PatternFill("solid",fgColor="FF4444")
    if val>=0.06:             return PatternFill("solid",fgColor="ED7D31")
    if val>=0.03:             return PatternFill("solid",fgColor="FFC000")
    if val>=0.01:             return PatternFill("solid",fgColor="92D050")
    return                         PatternFill("solid",fgColor="E2EFDA")

def issue_font(val):
    if val and val>=0.06:
        return Font(name="Calibri",size=9,bold=True,color="FFFFFF")
    return Font(name="Calibri",size=9)

def feat_fill(val):
    if not val:               return PatternFill("solid",fgColor="FFFFFF")
    if val>=10:               return PatternFill("solid",fgColor="1F4E79")
    if val>=5:                return PatternFill("solid",fgColor="2E75B6")
    if val>=2:                return PatternFill("solid",fgColor="BDD7EE")
    return                         PatternFill("solid",fgColor="DEEAF1")

def feat_font(val):
    if val and val>=5:
        return Font(name="Calibri",size=9,bold=True,color="FFFFFF")
    return Font(name="Calibri",size=9)

# ══════════════════════════════════════════════════════════════════════════════
#  EXCEL GENERATOR  — 5-sheet professional report
# ══════════════════════════════════════════════════════════════════════════════

def generate_excel(analysis,insights,app_name,months):
    wb=Workbook()
    monthly   =analysis["monthly"]
    issue_df  =analysis["issue_df"]
    feature_df=analysis["feature_df"]
    month_cols=issue_df.columns.tolist() if not issue_df.empty else monthly["month_str"].tolist()

    # ── Style constants ───────────────────────────────────────────────────────
    DB="1F4E79"; MB="2E75B6"; LB="D6E4F0"
    db_f=PatternFill("solid",fgColor=DB); mb_f=PatternFill("solid",fgColor=MB)
    lb_f=PatternFill("solid",fgColor=LB); alt_f=PatternFill("solid",fgColor="EBF3FB")
    wh_f=PatternFill("solid",fgColor="FFFFFF")
    H1=Font(name="Calibri",bold=True,size=14,color="FFFFFF")
    H2=Font(name="Calibri",bold=True,size=11,color="FFFFFF")
    H3=Font(name="Calibri",bold=True,size=10,color="FFFFFF")
    BD=Font(name="Calibri",size=10); BD9=Font(name="Calibri",size=9)
    CTR=Alignment(horizontal="center",vertical="center",wrap_text=True)
    LFT=Alignment(horizontal="left",  vertical="center",wrap_text=True)
    thin=Side(style="thin",color="BFBFBF")
    bdr=Border(left=thin,right=thin,top=thin,bottom=thin)

    def hdr_row(ws,ri,fill,fnt=H3):
        for c in ws[ri]:
            c.fill=fill; c.font=fnt; c.alignment=CTR; c.border=bdr

    # ══════════════════════════════════════════════════════════════════════════
    # S1 : Dashboard
    # ══════════════════════════════════════════════════════════════════════════
    ws1=wb.active; ws1.title="📊 Dashboard"
    ws1.column_dimensions["A"].width=14
    for col in list("BCDEFGHIJKLMNO"):
        ws1.column_dimensions[col].width=12

    ws1.merge_cells("A1:N1")
    ws1["A1"]=f"{app_name} — PlayStore Review Trend Analysis"
    ws1["A1"].font=H1; ws1["A1"].fill=db_f; ws1["A1"].alignment=CTR; ws1.row_dimensions[1].height=24

    ws1.merge_cells("A2:N2")
    ws1["A2"]=(f"Period: {month_cols[0]} – {month_cols[-1]}  |  Total Reviews: {analysis['total_reviews']:,}"
               f"  |  Avg Rating: {analysis['avg_rating']} ★  |  Generated: {datetime.now().strftime('%d %b %Y')}")
    ws1["A2"].font=Font(name="Calibri",size=10,color="FFFFFF")
    ws1["A2"].fill=PatternFill("solid",fgColor="2E4057"); ws1["A2"].alignment=CTR; ws1.row_dimensions[2].height=18

    # KPI tiles
    kpi_data=[
        ("Peak Month",f"{monthly.loc[monthly['count'].idxmax(),'month_str']} ({monthly['count'].max():,})","Highest review volume","2E75B6"),
        ("Avg Rating",f"{analysis['avg_rating']} / 5.0","Across all months","375623"),
        ("Best Rating Month",f"{monthly.loc[monthly['avg_rating'].idxmax(),'month_str']} ({monthly['avg_rating'].max():.2f}★)","Highest avg rating","7030A0"),
        ("Lowest Neg% Month",f"{monthly.loc[monthly['neg_pct'].idxmin(),'month_str']} ({monthly['neg_pct'].min():.1f}%)","Fewest 1+2★ reviews","C00000"),
    ]
    kpi_col_pairs=[("B","D"),("E","G"),("H","J"),("K","M")]
    for (sc,ec),(title,val,sub,color) in zip(kpi_col_pairs,kpi_data):
        f=PatternFill("solid",fgColor=color)
        for row,txt,fnt in [(4,title,Font(name="Calibri",bold=True,size=9,color="FFFFFF")),
                             (5,val,  Font(name="Calibri",bold=True,size=13,color="FFFFFF")),
                             (6,sub,  Font(name="Calibri",size=9,color=LB))]:
            ws1.merge_cells(f"{sc}{row}:{ec}{row}")
            ws1[f"{sc}{row}"]=txt; ws1[f"{sc}{row}"].fill=f
            ws1[f"{sc}{row}"].font=fnt; ws1[f"{sc}{row}"].alignment=CTR
        ws1.row_dimensions[5].height=26

    # Monthly table
    ws1.merge_cells("A8:J8")
    ws1["A8"]="Monthly Rating Summary"
    ws1["A8"].font=H2; ws1["A8"].fill=db_f; ws1["A8"].alignment=CTR; ws1.row_dimensions[8].height=18

    mhdrs=["Month","Total Reviews","★ Avg","1★","2★","3★","4★","5★","1+2★ %","4+5★ %"]
    for ci,h in enumerate(mhdrs,1):
        ws1.cell(9,ci,h).fill=mb_f; ws1.cell(9,ci).font=H3
        ws1.cell(9,ci).alignment=CTR; ws1.cell(9,ci).border=bdr

    for ri,(_, row) in enumerate(monthly.iterrows(),start=10):
        fill=alt_f if ri%2==0 else wh_f
        def safe_int(x):
            try: return int(x) if x is not None and not (isinstance(x,float) and np.isnan(x)) else 0
            except: return 0
        for ci,v in enumerate([row["month_str"],safe_int(row["count"]),row["avg_rating"],
                                safe_int(row["one_star"]),safe_int(row["two_star"]),safe_int(row["three_star"]),
                                safe_int(row["four_star"]),safe_int(row["five_star"]),
                                f'{row["neg_pct"]}%',f'{row["pos_pct"]}%'],1):
            c=ws1.cell(ri,ci,v); c.fill=fill; c.font=BD9; c.alignment=CTR; c.border=bdr
        ws1.row_dimensions[ri].height=15

    # ══════════════════════════════════════════════════════════════════════════
    # S2 : Issue Trends heatmap
    # ══════════════════════════════════════════════════════════════════════════
    ws2=wb.create_sheet("🔍 Issue Trends")
    ws2.column_dimensions["A"].width=30
    for i in range(2,len(month_cols)+2):
        ws2.column_dimensions[get_column_letter(i)].width=10
    nc=len(month_cols)
    ws2.column_dimensions[get_column_letter(nc+2)].width=24
    ws2.column_dimensions[get_column_letter(nc+3)].width=22
    ws2.column_dimensions[get_column_letter(nc+4)].width=16

    lc=get_column_letter(nc+4)
    ws2.merge_cells(f"A1:{lc}1")
    ws2["A1"]="Issue Trend Analysis — % of Negative (1+2★) Reviews Mentioning Each Issue"
    ws2["A1"].font=H1; ws2["A1"].fill=db_f; ws2["A1"].alignment=CTR; ws2.row_dimensions[1].height=20

    ws2.merge_cells(f"A2:{lc}2")
    ws2["A2"]="Values = % of that month's negative reviews. 🔴 ≥15% Critical  |  🟠 10-14% High  |  🟡 6-9% Medium  |  🟢 3-5% Low  |  ⚪ 1-2% Minimal  |  ⬜ 0% None"
    ws2["A2"].font=Font(name="Calibri",size=9,italic=True); ws2["A2"].fill=PatternFill("solid",fgColor="F2F2F2")
    ws2["A2"].alignment=LFT; ws2.row_dimensions[2].height=16

    ws2.cell(4,1,"Issue Category").fill=db_f; ws2.cell(4,1).font=H3
    ws2.cell(4,1).alignment=CTR; ws2.cell(4,1).border=bdr
    for ci,m in enumerate(month_cols,2):
        c=ws2.cell(4,ci,m); c.fill=mb_f; c.font=H3; c.alignment=CTR; c.border=bdr
    for ci,lbl in enumerate(["Trend","Peak Month","Status"],nc+2):
        c=ws2.cell(4,ci,lbl); c.fill=db_f; c.font=H3; c.alignment=CTR; c.border=bdr
    ws2.row_dimensions[4].height=18

    trend_fills={
        "✅ Improved":       PatternFill("solid",fgColor="E2EFDA"),
        "⚠ Persists":       PatternFill("solid",fgColor="FFF2CC"),
        "⚡ Recurring Spike":PatternFill("solid",fgColor="FCE4D6"),
        "🔺 Rising":        PatternFill("solid",fgColor="FDECEA"),
        "✅ Stable-Low":    PatternFill("solid",fgColor="F4FBF4"),
    }
    for ri,cat in enumerate(issue_df.index,start=5):
        vals=issue_df.loc[cat].tolist()
        trend=classify_trend(vals); pm=peak_month(vals,month_cols)
        status=trend.split(" ",1)[-1] if " " in trend else trend
        ws2.row_dimensions[ri].height=16
        cc=ws2.cell(ri,1,cat); cc.font=Font(name="Calibri",size=9,bold=True)
        cc.fill=lb_f if ri%2==0 else wh_f; cc.alignment=LFT; cc.border=bdr
        for ci,val in enumerate(vals,2):
            c=ws2.cell(ri,ci,f"{val*100:.0f}%" if val and val>0 else "")
            c.fill=issue_fill(val); c.font=issue_font(val); c.alignment=CTR; c.border=bdr
        tf=trend_fills.get(trend,wh_f)
        for ci,v2 in enumerate([trend,pm,status],nc+2):
            c=ws2.cell(ri,ci,v2); c.fill=tf; c.font=BD9; c.alignment=CTR; c.border=bdr

    lr=len(issue_df)+7
    ws2.cell(lr,1,"HEAT MAP LEGEND:").font=Font(name="Calibri",bold=True,size=9)
    for ci,(lbl,bg,fg) in enumerate([("≥15% Critical","C00000","FFFFFF"),("10–14% High","FF4444","FFFFFF"),
                                      ("6–9% Medium","ED7D31","FFFFFF"),("3–5% Low","FFC000","000000"),
                                      ("1–2% Minimal","92D050","000000"),("0% None","FFFFFF","000000")],1):
        c=ws2.cell(lr+1,ci,lbl); c.fill=PatternFill("solid",fgColor=bg)
        c.font=Font(name="Calibri",size=9,bold=True,color=fg); c.alignment=CTR; c.border=bdr

    # ══════════════════════════════════════════════════════════════════════════
    # S3 : Feature Requests
    # ══════════════════════════════════════════════════════════════════════════
    ws3=wb.create_sheet("💡 Feature Requests")
    ws3.column_dimensions["A"].width=40
    for i in range(2,len(month_cols)+2):
        ws3.column_dimensions[get_column_letter(i)].width=10
    tc=get_column_letter(nc+2); pc=get_column_letter(nc+3)
    ws3.column_dimensions[tc].width=10; ws3.column_dimensions[pc].width=14

    ws3.merge_cells(f"A1:{pc}1")
    ws3["A1"]="User Feature Requests — Mentions per Month (from Review Text Analysis)"
    ws3["A1"].font=H1; ws3["A1"].fill=db_f; ws3["A1"].alignment=CTR
    ws3.merge_cells(f"A2:{pc}2")
    ws3["A2"]="Count of reviews mentioning each feature. Blank = 0. Highlighted = notable demand."
    ws3["A2"].font=Font(name="Calibri",size=9,italic=True)
    ws3["A2"].fill=PatternFill("solid",fgColor="F2F2F2"); ws3["A2"].alignment=LFT

    ws3.cell(4,1,"Feature Request").fill=db_f; ws3.cell(4,1).font=H3
    ws3.cell(4,1).alignment=CTR; ws3.cell(4,1).border=bdr
    for ci,m in enumerate(month_cols,2):
        c=ws3.cell(4,ci,m); c.fill=mb_f; c.font=H3; c.alignment=CTR; c.border=bdr
    for ci2,lbl in enumerate(["Total","Priority"],nc+2):
        c=ws3.cell(4,ci2,lbl); c.fill=db_f; c.font=H3; c.alignment=CTR; c.border=bdr
    ws3.row_dimensions[4].height=18

    if not feature_df.empty:
        feat_totals=feature_df.fillna(0).sum(axis=1).sort_values(ascending=False)
        ri=5
        for feat in feat_totals.index:
            total_val=int(feat_totals[feat])
            if total_val==0:
                continue
            vals=[feature_df.loc[feat,m] if m in feature_df.columns else None for m in month_cols]
            priority="🔴 High" if total_val>=20 else ("🟡 Medium" if total_val>=8 else "🟢 Low")
            ws3.row_dimensions[ri].height=16
            fc=ws3.cell(ri,1,feat); fc.font=Font(name="Calibri",size=9,bold=True)
            fc.fill=lb_f if ri%2==0 else wh_f; fc.alignment=LFT; fc.border=bdr
            for ci,val in enumerate(vals,2):
                safe_val = int(val) if (val is not None and isinstance(val,(int,float)) and not np.isnan(val) and val>0) else None
                c=ws3.cell(ri,ci,safe_val if safe_val else "")
                c.fill=feat_fill(safe_val); c.font=feat_font(safe_val); c.alignment=CTR; c.border=bdr
            # Total
            tcc=ws3.cell(ri,nc+2,total_val)
            tcc.fill=mb_f if total_val>=20 else (PatternFill("solid",fgColor="BDD7EE") if total_val>=8 else wh_f)
            tcc.font=Font(name="Calibri",size=9,bold=True,color="FFFFFF" if total_val>=20 else "000000")
            tcc.alignment=CTR; tcc.border=bdr
            # Priority
            pcc=ws3.cell(ri,nc+3,priority)
            pcc.fill=(PatternFill("solid",fgColor="C00000") if total_val>=20
                      else PatternFill("solid",fgColor="FFC000") if total_val>=8
                      else PatternFill("solid",fgColor="E2EFDA"))
            pcc.font=Font(name="Calibri",size=9,bold=True,color="FFFFFF" if total_val>=20 else "000000")
            pcc.alignment=CTR; pcc.border=bdr
            ri+=1
    else:
        ws3.cell(5,1,"No feature request keywords detected for this period.")
        ws3.cell(5,1).font=Font(name="Calibri",size=10,italic=True,color="808080")

    # ══════════════════════════════════════════════════════════════════════════
    # S4 : Key Insights
    # ══════════════════════════════════════════════════════════════════════════
    ws4=wb.create_sheet("📝 Key Insights")
    ws4.column_dimensions["A"].width=28
    ws4.column_dimensions["B"].width=42
    ws4.column_dimensions["C"].width=65

    ws4.merge_cells("A1:C1")
    ws4["A1"]=f"{app_name} — Key Trend Insights ({month_cols[0]} – {month_cols[-1]})"
    ws4["A1"].font=H1; ws4["A1"].fill=db_f; ws4["A1"].alignment=CTR; ws4.row_dimensions[1].height=22

    for ci,(h,w) in enumerate(zip(["Area","Finding","Supporting Detail"],["A","B","C"]),1):
        ws4.cell(3,ci,h).fill=mb_f; ws4.cell(3,ci).font=H2
        ws4.cell(3,ci).alignment=CTR; ws4.cell(3,ci).border=bdr

    row=4
    def section_hdr(ws,r,title,color):
        ws.merge_cells(f"A{r}:C{r}")
        ws[f"A{r}"]=title
        ws[f"A{r}"].fill=PatternFill("solid",fgColor=color)
        ws[f"A{r}"].font=Font(name="Calibri",bold=True,size=10,color="FFFFFF")
        ws[f"A{r}"].alignment=LFT; ws[f"A{r}"].border=bdr
        ws.row_dimensions[r].height=18

    def ins_row(ws,r,area,finding,detail,rfill):
        for ci,v in enumerate([area,finding,detail],1):
            c=ws.cell(r,ci,v); c.fill=rfill; c.border=bdr
            c.font=Font(name="Calibri",bold=(ci==1),size=9)
            c.alignment=LFT
        ws.row_dimensions[r].height=34

    if insights["improved"]:
        section_hdr(ws4,row,"✅  WHAT IMPROVED  —  Issues that declined significantly over the period","375623"); row+=1
        for it in insights["improved"]:
            ins_row(ws4,row,it["issue"],it["finding"],it["detail"],PatternFill("solid",fgColor="E2EFDA") if row%2==0 else wh_f); row+=1
        row+=1

    if insights["persistent"]:
        section_hdr(ws4,row,"⚠  WHAT STAYED CONSTANT  —  Persistent issues with no clear resolution","7030A0"); row+=1
        for it in insights["persistent"]:
            ins_row(ws4,row,it["issue"],it["finding"],it["detail"],PatternFill("solid",fgColor="EDE7F6") if row%2==0 else wh_f); row+=1
        row+=1

    spike_list=insights["spikes"]+insights["rising"]
    if spike_list:
        section_hdr(ws4,row,"⚡  NEW & RECURRING SPIKES  —  Issues that appeared or intensified","C00000"); row+=1
        for it in spike_list:
            ins_row(ws4,row,it["issue"],it["finding"],it["detail"],PatternFill("solid",fgColor="FCE4D6") if row%2==0 else wh_f); row+=1
        row+=1

    if insights["feature_requests"]:
        section_hdr(ws4,row,"💡  TOP FEATURE REQUESTS  —  Consistent asks from users","ED7D31"); row+=1
        for it in insights["feature_requests"][:12]:
            detail=f"Total {it['total']} mentions. Users consistently request this — likely a competitive gap vs rival apps."
            ins_row(ws4,row,it["feature"],it["finding"],detail,PatternFill("solid",fgColor="FFF2CC") if row%2==0 else wh_f); row+=1
        row+=1

    section_hdr(ws4,row,"🎯  STRATEGIC SUMMARY",DB); row+=1
    for it in insights["strategic"]:
        ins_row(ws4,row,it["area"],it["finding"],it["detail"],lb_f if row%2==0 else wh_f); row+=1

    # ══════════════════════════════════════════════════════════════════════════
    # S5 : All Reviews
    # ══════════════════════════════════════════════════════════════════════════
    ws5=wb.create_sheet("📋 All Reviews")
    df_all=analysis["df"].copy()
    SENT_FILLS={"Positive":PatternFill("solid",fgColor="E2EFDA"),
                "Negative":PatternFill("solid",fgColor="FCE4D6"),
                "Neutral": PatternFill("solid",fgColor="F5F5F5")}
    TAG_FILL = PatternFill("solid",fgColor="FFF2CC")   # light amber for tag cells

    # issue_tags column already on df from analyse_reviews; include it only if present
    base_cols={"userName":"Reviewer","score":"Rating","content":"Review Text",
               "at":"Date","thumbsUpCount":"Helpful","replyContent":"Dev Reply","sentiment":"Sentiment"}
    avail=[c for c in base_cols if c in df_all.columns]

    # Add issue_tags as the last column if it exists and has at least one non-empty value
    has_tags = "issue_tags" in df_all.columns and df_all["issue_tags"].str.strip().astype(bool).any()

    headers = [base_cols[c] for c in avail] + (["Issue Tags"] if has_tags else [])
    ws5.append(headers)
    hdr_row(ws5,1,mb_f,H3)

    cw={"Reviewer":18,"Rating":8,"Review Text":55,"Date":12,"Helpful":12,
        "Dev Reply":40,"Sentiment":12,"Issue Tags":40}
    for i,h in enumerate(headers,1):
        ws5.column_dimensions[get_column_letter(i)].width=cw.get(h,15)

    disp=df_all[avail].copy()
    disp["at"]=pd.to_datetime(disp["at"]).dt.strftime("%Y-%m-%d")

    tag_col_idx = len(avail) + 1  # 1-based index of Issue Tags column

    for idx,rv in enumerate(disp.itertuples(index=False)):
        row_vals = list(rv)
        if has_tags:
            row_vals.append(df_all["issue_tags"].iloc[idx])
        ws5.append(row_vals)
        ri=ws5.max_row
        sent=rv._asdict().get("sentiment","Neutral")
        base_fill=SENT_FILLS.get(sent,SENT_FILLS["Neutral"])
        for ci,cell in enumerate(ws5[ri],1):
            if has_tags and ci == tag_col_idx:
                # Only colour tag cell if it actually has a tag
                cell.fill = TAG_FILL if cell.value else PatternFill("solid",fgColor="FFFFFF")
                cell.font = Font(name="Calibri",size=8,italic=True,
                                 color="7B5F00" if cell.value else "AAAAAA")
            else:
                cell.fill=base_fill; cell.font=BD9
            cell.alignment=Alignment(horizontal="left",vertical="center",wrap_text=True)
            cell.border=bdr
        ws5.row_dimensions[ri].height=35
    ws5.freeze_panes="A2"; ws5.auto_filter.ref=ws5.dimensions

    buf=io.BytesIO(); wb.save(buf); return buf.getvalue()

# ══════════════════════════════════════════════════════════════════════════════
#  STREAMLIT UI
# ══════════════════════════════════════════════════════════════════════════════

st.set_page_config(page_title="Play Store Review Analyser",page_icon="📱",
                   layout="wide",initial_sidebar_state="expanded")

st.markdown("""
<style>
[data-testid="stSidebar"]{background:linear-gradient(180deg,#1F4E79 0%,#154069 100%);}
[data-testid="stSidebar"] p,[data-testid="stSidebar"] span,
[data-testid="stSidebar"] div,[data-testid="stSidebar"] label{color:#E8F1FB !important;}
.stButton>button{background:linear-gradient(135deg,#1F4E79,#2E75B6);color:white;border:none;
  border-radius:8px;padding:8px 18px;font-weight:600;transition:all 0.2s ease;}
.stButton>button:hover{background:linear-gradient(135deg,#2E75B6,#1F4E79);
  box-shadow:0 4px 12px rgba(31,78,121,0.35);transform:translateY(-1px);}
[data-testid="stDownloadButton"]>button{background:linear-gradient(135deg,#375623,#70AD47) !important;
  color:white !important;border:none !important;border-radius:8px !important;font-weight:600 !important;}
[data-testid="stMetric"]{background:linear-gradient(135deg,#EBF3FB,#D6E4F0);border-radius:10px;
  padding:12px 16px;border-left:4px solid #2E75B6;box-shadow:0 1px 4px rgba(0,0,0,0.08);}
.tag{display:inline-block;padding:4px 11px;border-radius:14px;font-size:13px;margin:3px;font-weight:500;}
.tag-pos{background:#E2EFDA;color:#375623;border:1px solid #A8D08D;}
.tag-neg{background:#FCE4D6;color:#C00000;border:1px solid #F4B183;}
.tag-rise{background:#FCE4D6;color:#C00000;border:1px solid #F4B183;}
.tag-fall{background:#E2EFDA;color:#375623;border:1px solid #A8D08D;}
.tag-pers{background:#EDE7F6;color:#7030A0;border:1px solid #C5A8E0;}
[data-testid="stSidebar"] .stButton>button{background:rgba(255,255,255,0.12) !important;
  border:1px solid rgba(255,255,255,0.25) !important;border-radius:8px !important;
  text-align:left !important;padding:6px 10px !important;font-size:13px !important;
  line-height:1.4 !important;color:#E8F1FB !important;margin-bottom:2px;}
[data-testid="stSidebar"] .stButton>button:hover{background:rgba(255,255,255,0.22) !important;transform:none;}
</style>
""",unsafe_allow_html=True)

for k in ["prefill_url","results","excel_cache","cache_key"]:
    if k not in st.session_state:
        st.session_state[k]=None

with st.sidebar:
    st.markdown("## 📱 Review Analyser"); st.markdown("---")
    if st.button("➕  New Search",use_container_width=True):
        for k in ["prefill_url","results","excel_cache","cache_key"]:
            st.session_state[k]=None
        st.rerun()
    st.markdown("### 🕑 History")
    history=load_history()
    if not history:
        st.caption("No history yet.")
    else:
        for item in history:
            c1,c2=st.columns([5,1])
            with c1:
                if st.button(f"{item['app_name'][:24]}\n🗓 {item['last_searched'][:10]}",
                             key=f"h_{item['app_id']}",use_container_width=True):
                    st.session_state.prefill_url=item["url"]
                    st.session_state.results=None; st.session_state.excel_cache=None; st.rerun()
            with c2:
                if st.button("✕",key=f"d_{item['app_id']}",help="Remove"):
                    history=[h for h in history if h["app_id"]!=item["app_id"]]
                    save_history(history); st.rerun()

st.markdown('<h1 style="color:#1F4E79;margin-bottom:2px">📱 Play Store Review Analyser</h1>',unsafe_allow_html=True)
st.caption("Paste any Play Store URL · choose a time period · get a full-depth Excel report with issue heatmaps, feature requests & key insights.")
st.markdown("---")

c_url,c_dur,c_btn=st.columns([4,2,1])
with c_url:
    url_input=st.text_input("🔗 Play Store App URL",value=st.session_state.prefill_url or "",
                             placeholder="https://play.google.com/store/apps/details?id=com.example.app")
with c_dur:
    dur_label=st.selectbox("📅 Time Period",list(MONTHS_OPTIONS.keys()),index=3)
    months=MONTHS_OPTIONS[dur_label]
with c_btn:
    st.markdown("<br>",unsafe_allow_html=True)
    get_data=st.button("🔍  Get Data",type="primary",use_container_width=True)

if get_data:
    if not url_input.strip():
        st.warning("Please enter a Play Store URL.")
    else:
        app_id=extract_app_id(url_input.strip())
        if not app_id:
            st.error("⚠️ Couldn't find an app ID in the URL.")
        else:
            with st.spinner(f"Fetching & analysing reviews for the last {months} month(s)…"):
                try:
                    app_info=fetch_app_info(app_id)
                    raw=fetch_reviews_cached(app_id,months)
                    if not raw:
                        st.warning("No reviews found for this period.")
                    else:
                        df=pd.DataFrame(raw)
                        analysis=analyse_reviews(df)
                        app_name=app_info.get("title",app_id)
                        ins=generate_insights(analysis["issue_df"],analysis["feature_df"],
                                              analysis["monthly"],app_name)
                        add_to_history(app_id,app_name,url_input.strip(),app_info.get("icon",""))
                        st.session_state.results=dict(analysis=analysis,insights=ins,
                            app_name=app_name,app_info=app_info,months=months,
                            dur_label=dur_label,app_id=app_id)
                        st.session_state.excel_cache=None
                        st.session_state.cache_key=f"{app_id}_{months}"
                        st.rerun()
                except Exception as e:
                    st.error(f"❌ Error: {e}")

if st.session_state.results:
    res=st.session_state.results
    analysis=res["analysis"]; insights=res["insights"]
    app_name=res["app_name"]; app_info=res["app_info"]
    months=res["months"]; app_id=res["app_id"]

    ic,inf=st.columns([1,9])
    with ic:
        if app_info.get("icon"): st.image(app_info["icon"],width=60)
    with inf:
        st.markdown(f"### {app_name}")
        st.caption(f"**Developer:** {app_info.get('developer','—')}  ·  "
                   f"**Category:** {app_info.get('genre','—')}  ·  "
                   f"**Overall Rating:** {app_info.get('score','—')} ⭐  ·  Period: {res['dur_label']}")
    st.markdown("---")

    total=analysis["total_reviews"]
    pos=int(analysis["sentiment_dist"].get("Positive",0))
    neg=int(analysis["sentiment_dist"].get("Negative",0))
    k1,k2,k3,k4,k5,k6=st.columns(6)
    k1.metric("📊 Total Reviews",f"{total:,}")
    k2.metric("⭐ Avg Rating",analysis["avg_rating"])
    k3.metric("📅 Recent 30-Day",analysis["recent_avg"])
    k4.metric("😊 Positive",f"{pos} ({round(pos/max(total,1)*100,1)}%)")
    k5.metric("😞 Negative",f"{neg} ({round(neg/max(total,1)*100,1)}%)")
    k6.metric("💬 Dev Reply Rate",f"{analysis['reply_rate']}%")

    st.markdown("<br>",unsafe_allow_html=True)

    tab1,tab2,tab3,tab4,tab5=st.tabs(["📈 Charts","🔍 Issue Tracker","💡 Feature Requests","💬 Themes","📋 Raw Data"])

    with tab1:
        cl,cr=st.columns(2)
        with cl:
            rd=analysis["rating_dist"].reset_index(); rd.columns=["Stars","Count"]
            rd["Stars"]=rd["Stars"].astype(str)+" ⭐"
            fig=px.bar(rd,x="Stars",y="Count",color="Count",text="Count",
                       color_continuous_scale=["#C00000","#ED7D31","#FFC000","#70AD47","#375623"],
                       title="Rating Distribution")
            fig.update_traces(textposition="outside")
            fig.update_layout(showlegend=False,coloraxis_showscale=False,
                              plot_bgcolor="white",paper_bgcolor="white",margin=dict(t=40,b=10,l=10,r=10))
            st.plotly_chart(fig,use_container_width=True)
        with cr:
            sd=analysis["sentiment_dist"].reset_index(); sd.columns=["Sentiment","Count"]
            fig2=px.pie(sd,values="Count",names="Sentiment",hole=0.42,title="Sentiment Distribution",
                        color="Sentiment",color_discrete_map={"Positive":"#375623","Negative":"#C00000","Neutral":"#FFC000"})
            fig2.update_layout(plot_bgcolor="white",paper_bgcolor="white",margin=dict(t=40,b=10,l=10,r=10))
            st.plotly_chart(fig2,use_container_width=True)
        if len(analysis["monthly"])>1:
            m=analysis["monthly"]
            fig3=go.Figure()
            fig3.add_trace(go.Bar(name="Positive",x=m["month_str"],y=m["positive"],marker_color="#375623"))
            fig3.add_trace(go.Bar(name="Neutral", x=m["month_str"],y=m["neutral"], marker_color="#FFC000"))
            fig3.add_trace(go.Bar(name="Negative",x=m["month_str"],y=m["negative"],marker_color="#C00000"))
            fig3.add_trace(go.Scatter(name="Avg Rating",x=m["month_str"],y=m["avg_rating"],
                                       yaxis="y2",mode="lines+markers",
                                       line=dict(color="#1F4E79",width=2.5),marker=dict(size=6)))
            fig3.update_layout(barmode="stack",title="Monthly Volume & Sentiment",
                                yaxis=dict(title="Review Count"),
                                yaxis2=dict(title="Avg Rating",overlaying="y",side="right",range=[0,5.5]),
                                plot_bgcolor="white",paper_bgcolor="white",
                                legend=dict(orientation="h",yanchor="bottom",y=1.02),
                                margin=dict(t=60,b=10,l=10,r=10))
            st.plotly_chart(fig3,use_container_width=True)

    with tab2:
        c1,c2,c3=st.columns(3)
        with c1:
            st.markdown("#### 🔺 Rising Issues")
            for i in insights["rising"]:
                st.markdown(f'<span class="tag tag-rise">🔺 {i["issue"]}</span>',unsafe_allow_html=True)
                st.caption(i["finding"])
            if not insights["rising"]: st.info("None detected.")
        with c2:
            st.markdown("#### ✅ Resolving Issues")
            for i in insights["improved"]:
                st.markdown(f'<span class="tag tag-fall">✅ {i["issue"]}</span>',unsafe_allow_html=True)
                st.caption(i["finding"])
            if not insights["improved"]: st.info("None detected.")
        with c3:
            st.markdown("#### ⚠️ Persistent Issues")
            for i in insights["persistent"]:
                st.markdown(f'<span class="tag tag-pers">⚠️ {i["issue"]}</span>',unsafe_allow_html=True)
                st.caption(i["finding"])
            if not insights["persistent"]: st.info("None detected.")
        if not analysis["issue_df"].empty:
            st.markdown("---")
            st.markdown("#### Issue Heatmap (% of negative reviews)")
            disp_i=analysis["issue_df"].copy()
            disp_i=disp_i[disp_i.max(axis=1)>=0.01]
            st.dataframe(disp_i.applymap(lambda x:f"{x*100:.0f}%" if x and x>0 else ""),
                         use_container_width=True,height=420)

    with tab3:
        feat_df=analysis["feature_df"]
        if not feat_df.empty:
            st.markdown("#### 💡 Feature Requests by Month")
            st.caption("Count of reviews mentioning each feature.")
            ft=feat_df.fillna(0).sum(axis=1).sort_values(ascending=False)
            fd=feat_df.loc[ft[ft>0].index].copy()
            fd.insert(0,"Total",ft[fd.index].astype(int))
            st.dataframe(fd.fillna(""),use_container_width=True,height=400)
        else:
            st.info("No feature request keywords found for this period.")

    with tab4:
        cp,cn=st.columns(2)
        with cp:
            st.markdown("#### 😊 What Happy Users Love")
            for kw,cnt in analysis["positive_themes"][:12]:
                st.markdown(f'<span class="tag tag-pos">👍 {kw} ({cnt})</span>',unsafe_allow_html=True)
            st.markdown("##### Top Positive Reviews")
            for _,row in analysis["best_reviews"].head(3).iterrows():
                st.markdown(f"> *\"{str(row['content'])[:200]}...\"*")
                st.caption(f"{'⭐'*int(row['score'])}  ·  {str(row['at'])[:10]}  ·  👍 {int(row.get('thumbsUpCount',0))}")
                st.markdown("---")
        with cn:
            st.markdown("#### 😞 What Critics Complain About")
            for kw,cnt in analysis["negative_themes"][:12]:
                st.markdown(f'<span class="tag tag-neg">👎 {kw} ({cnt})</span>',unsafe_allow_html=True)
            st.markdown("##### Top Critical Reviews")
            for _,row in analysis["worst_reviews"].head(3).iterrows():
                st.markdown(f"> *\"{str(row['content'])[:200]}...\"*")
                st.caption(f"{'⭐'*int(row['score'])}  ·  {str(row['at'])[:10]}")
                st.markdown("---")

    with tab5:
        dc=["userName","score","sentiment","content","at","thumbsUpCount"]
        disp=analysis["df"][[c for c in dc if c in analysis["df"].columns]].copy()
        disp["at"]=pd.to_datetime(disp["at"]).dt.strftime("%Y-%m-%d")
        disp.columns=[c.replace("userName","Reviewer").replace("score","Rating")
                       .replace("sentiment","Sentiment").replace("content","Review")
                       .replace("at","Date").replace("thumbsUpCount","Helpful") for c in disp.columns]
        st.dataframe(disp,use_container_width=True,height=500)

    st.markdown("---")
    dl1,_=st.columns([2,6])
    cache_key=f"{app_id}_{months}"
    if st.session_state.cache_key!=cache_key or st.session_state.excel_cache is None:
        with st.spinner("Building Excel report…"):
            st.session_state.excel_cache=generate_excel(analysis,insights,app_name,months)
        st.session_state.cache_key=cache_key
    with dl1:
        st.download_button("📥 Download Full Excel Report",
                           data=st.session_state.excel_cache,
                           file_name=f"{app_id}_{months}mo_analysis.xlsx",
                           mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                           use_container_width=True)
