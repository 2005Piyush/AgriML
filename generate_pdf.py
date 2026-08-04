"""
AgriML — PDF Presentation Generator
Builds a 22-page A4-landscape PDF with all diagrams/flowcharts embedded.
Run:  python generate_pdf.py
"""
import os, base64, subprocess, sys

BASE     = os.path.dirname(os.path.abspath(__file__))
OUT_HTML = os.path.join(BASE, "AgriML_Presentation_Print.html")
OUT_PDF  = os.path.join(BASE, "AgriML_Presentation.pdf")

# ── Chrome paths ─────────────────────────────────────────────────────────────
CHROME_PATHS = [
    r"C:\Program Files\Google\Chrome\Application\chrome.exe",
    r"C:\Program Files (x86)\Google\Chrome\Application\chrome.exe",
    r"C:\Users\krpiy\AppData\Local\Google\Chrome\Application\chrome.exe",
]

def find_chrome():
    for p in CHROME_PATHS:
        if os.path.exists(p):
            return p
    return None

# ── Image loader (base64) ─────────────────────────────────────────────────────
def b64(fn):
    p = os.path.join(BASE, fn)
    if not os.path.exists(p):
        print(f"[WARN] Missing image: {fn}")
        return ""
    with open(p, "rb") as f:
        return "data:image/png;base64," + base64.b64encode(f.read()).decode()

I = {k: b64(v) for k, v in {
    "fw": "AgriML_Model_Flowchart.png",
    "dp": "data_pipeline_flow.png",
    "rf": "random_forest_diagram.png",
    "gb": "gradient_boosting_diagram.png",
    "ce": "climate_engine_diagram.png",
    "rr": "risk_prediction_radar.png",
    "sh": "soil_health_trend.png",
    "al": "adaptive_learning_loop.png",
}.items()}

# ── CSS ───────────────────────────────────────────────────────────────────────
CSS = """
@import url('https://fonts.googleapis.com/css2?family=Inter:wght@300;400;600;700;800;900&display=swap');
*{margin:0;padding:0;box-sizing:border-box;}
body{font-family:'Inter',sans-serif;background:#0F172A;color:#fff;}
.slide{
  width:297mm;height:210mm;
  padding:11mm 16mm 8mm;
  background:#0F172A;
  position:relative;overflow:hidden;
  display:flex;flex-direction:column;
  page-break-after:always;break-after:page;
}
.slide:last-child{page-break-after:avoid;break-after:avoid;}
.deco{position:absolute;border-radius:50%;opacity:.10;pointer-events:none;}
.slide-title{font-size:22pt;font-weight:800;color:#fff;line-height:1.15;}
.bar{width:60px;height:3px;background:#00C97B;border-radius:2px;margin:4px 0 5px;}
.subtitle{font-size:9.5pt;color:#8B9CB2;margin-bottom:10px;}
.card{background:#1A253C;border-radius:9px;padding:12px;border:1.5px solid transparent;}
.g{border-color:#00C97B;} .b{border-color:#38BDF8;} .o{border-color:#F5A623;}
.p{border-color:#E84D8A;} .pu{border-color:#A78BFA;}
.row{display:flex;gap:12px;flex:1;}
.col{flex:1;display:flex;flex-direction:column;}
.grid4{display:grid;grid-template-columns:repeat(4,1fr);gap:12px;flex:1;}
.grid3{display:grid;grid-template-columns:repeat(3,1fr);gap:12px;flex:1;}
.grid5{display:grid;grid-template-columns:repeat(5,1fr);gap:8px;flex:1;}
.grid6{display:grid;grid-template-columns:repeat(6,1fr);gap:7px;flex:1;}
ul.bul{list-style:none;}
ul.bul li{font-size:8.5pt;color:#CBD5E1;padding:2px 0;line-height:1.55;}
ul.bul li::before{content:"→ ";color:#00C97B;font-weight:700;}
table{width:100%;border-collapse:collapse;font-size:8pt;margin-top:4px;}
th{background:#00C97B;color:#0F172A;padding:6px 8px;text-align:center;font-weight:700;}
td{padding:5px 8px;text-align:center;border-bottom:1px solid #1e2d4a;color:#CBD5E1;}
tr:nth-child(even) td{background:#1A253C;}
tr:nth-child(odd) td{background:#141F33;}
img.diag{width:100%;height:100%;object-fit:contain;border-radius:8px;}
img.full{width:100%;max-height:105mm;object-fit:contain;border-radius:8px;}
.sct{font-size:10.5pt;font-weight:700;margin-bottom:6px;}
.stat{font-size:20pt;font-weight:900;color:#00C97B;}
.stat-lbl{font-size:8pt;color:#CBD5E1;}
.pg{position:absolute;bottom:5mm;right:14mm;font-size:7.5pt;color:#8B9CB2;}
.logo{position:absolute;bottom:5mm;left:14mm;font-size:7.5pt;color:#00C97B;font-weight:700;}
@page{size:A4 landscape;margin:0;}
@media print{
  *{-webkit-print-color-adjust:exact!important;print-color-adjust:exact!important;color-adjust:exact!important;}
}
"""

# ── Slide helpers ──────────────────────────────────────────────────────────────
def slide(body, n):
    return (
        f'<div class="slide">'
        f'<span class="logo">🌾 AgriML</span>'
        f'{body}'
        f'<span class="pg">{n} / 22</span>'
        f'</div>\n'
    )

def hdr(title, sub=None, bar="#00C97B"):
    s  = f'<div class="slide-title">{title}</div>'
    s += f'<div class="bar" style="background:{bar}"></div>'
    if sub:
        s += f'<div class="subtitle">{sub}</div>'
    return s

def img(key, cls="diag"):
    src = I.get(key, "")
    if not src:
        return f'<div style="display:flex;align-items:center;justify-content:center;height:100%;color:#8B9CB2;font-size:9pt;">[Image not found]</div>'
    return f'<img class="{cls}" src="{src}" alt="">'

def bullets(items, color="#00C97B"):
    li = "".join(f'<li style="--c:{color}">{x}</li>' for x in items)
    return f'<ul class="bul">{li}</ul>'

def table_html(headers, rows):
    ths = "".join(f"<th>{h}</th>" for h in headers)
    trs = ""
    for r in rows:
        trs += "<tr>" + "".join(f"<td>{c}</td>" for c in r) + "</tr>"
    return f"<table><tr>{ths}</tr>{trs}</table>"

def card(content, cls="card g", style=""):
    return f'<div class="card {cls}" style="{style}">{content}</div>'

# ══════════════════════════════════════════════════════════════════════════════
# BUILD ALL 22 SLIDES
# ══════════════════════════════════════════════════════════════════════════════
slides = []

# ── S1: Title ─────────────────────────────────────────────────────────────────
slides.append(slide(f"""
  <div class="deco" style="width:320px;height:320px;background:#00C97B;top:-110px;right:-90px;"></div>
  <div class="deco" style="width:240px;height:240px;background:#38BDF8;bottom:-80px;left:-70px;"></div>
  <div style="flex:1;display:flex;flex-direction:column;justify-content:center;">
    <div style="font-size:38pt;font-weight:900;color:#00C97B;margin-bottom:8px;">🌾 AgriML</div>
    <div style="font-size:19pt;font-weight:700;line-height:1.2;">AI-Powered Fertilizer Recommendation System</div>
    <div style="font-size:10pt;color:#8B9CB2;margin-top:10px;">Machine Learning  •  Climate-Aware  •  Sustainable Farming</div>
    <div style="font-size:9.5pt;color:#38BDF8;margin-top:20px;">Client-Side  •  Browser-Based  •  Zero Server Dependencies</div>
  </div>
""", 1))

# ── S2: Problem ───────────────────────────────────────────────────────────────
prob_cards = "".join([
    f'<div class="card o" style="text-align:center;padding:10px;display:flex;flex-direction:column;justify-content:center;">'
    f'<div class="stat">{v}</div><div class="stat-lbl">{l}</div></div>'
    for v, l in [("30%","of global fertilizer wasted"),("₹8,000 Cr","annual loss to Indian farmers"),
                 ("10 Crops","covered by AgriML"),("7 Types","fertilizer classes")]
])
slides.append(slide(f"""
  {hdr("The Problem","Why do farmers need intelligent fertilizer recommendations?","#F5A623")}
  <div class="row">
    <div class="col card b" style="padding:12px;">
      <div class="sct" style="color:#38BDF8;">Key Challenges</div>
      {bullets(["Over-fertilization wastes money and pollutes groundwater",
               "Under-fertilization leads to poor crop yields and food insecurity",
               "Nutrient imbalance — excess of one nutrient blocks others",
               "Climate variability changes optimal timing &amp; quantity",
               "Manual soil testing is expensive for small-hold farmers",
               "No easy agronomist access in rural India"])}
    </div>
    <div style="width:170px;display:flex;flex-direction:column;gap:8px;">{prob_cards}</div>
  </div>
""", 2))

# ── S3: Solution Overview ─────────────────────────────────────────────────────
sol_cards = ""
for icon, title, desc, cls, clr in [
    ("🧪","Data Input","Soil NPK, pH, moisture,\ntemperature, humidity, crop type","g","#00C97B"),
    ("⚙️","ML Engine","Random Forest Classifier\n+ Gradient Boosting Regressor","b","#38BDF8"),
    ("📊","Optimization","Quantity, cost &amp; environmental\nimpact scoring","o","#F5A623"),
    ("🌦️","Climate Aware","7-day forecast adjusts timing\nand quantity in real-time","p","#E84D8A"),
]:
    sol_cards += f'<div class="card {cls}" style="text-align:center;display:flex;flex-direction:column;align-items:center;justify-content:center;padding:14px;"><div style="font-size:26pt;">{icon}</div><div class="sct" style="color:{clr};margin:6px 0;">{title}</div><div style="font-size:8pt;color:#CBD5E1;">{desc}</div></div>'

slides.append(slide(f"""
  {hdr("Our Solution: AgriML","An end-to-end ML pipeline running entirely in the browser")}
  <div class="grid4">{sol_cards}</div>
  <div style="text-align:center;font-size:8pt;color:#8B9CB2;margin-top:8px;">→ Input Soil Data &nbsp;⟶&nbsp; ML Models Predict &nbsp;⟶&nbsp; Optimize Qty &amp; Cost &nbsp;⟶&nbsp; Climate-Adjusted Output</div>
""", 3))

# ── S4: System Architecture ───────────────────────────────────────────────────
layer_items = {
    "Data Layer": ("#00C97B", ["dataset.js — 2,000 synthetic samples","preprocessing.js — imputation + normalization","fertilizerModel.js — Random Forest","yieldModel.js — Gradient Boosting"]),
    "Analysis Layer": ("#F5A623", ["climateEngine.js — 7-day forecast","optimizer.js — qty, cost &amp; env","riskPredictor.js — 7 risk categories","sustainability.js — 6-component index"]),
    "Learning Layer": ("#A78BFA", ["adaptiveLearning.js — correction factors","feedbackLoop.js — farmer feedback","soilHealth.js — degradation tracker"]),
}
left_content = ""
for name, (clr, items) in layer_items.items():
    left_content += f'<div class="sct" style="color:{clr};margin-top:8px;">{name}</div>'
    left_content += bullets(items)

slides.append(slide(f"""
  {hdr("System Architecture","12 interconnected modules — 100% client-side JavaScript")}
  <div class="row">
    <div class="col card g" style="padding:12px;max-width:47%;">{left_content}</div>
    <div class="col" style="flex:1;">{img("fw")}</div>
  </div>
""", 4))

# ── S5: Data Pipeline ─────────────────────────────────────────────────────────
step_cards = ""
for num, title, desc, cls, clr in [
    ("1️⃣","Synthetic Data","2,000 samples\n10 crops × 7 fertilizers","g","#00C97B"),
    ("2️⃣","Imputation","Mean/Mode\nfill missing values","b","#38BDF8"),
    ("3️⃣","Normalization","Min-Max scaling\nto [0,1] range","o","#F5A623"),
    ("4️⃣","One-Hot Encode","10 binary columns\nfor crop type","p","#E84D8A"),
    ("5️⃣","Feature Vector","18-D combined\nvector → model","pu","#A78BFA"),
]:
    step_cards += f'<div class="card {cls}" style="text-align:center;padding:10px;display:flex;flex-direction:column;align-items:center;"><div style="font-size:16pt;">{num}</div><div style="font-size:9pt;font-weight:700;color:{clr};margin:4px 0;">{title}</div><div style="font-size:7.5pt;color:#CBD5E1;">{desc}</div></div>'

slides.append(slide(f"""
  {hdr("Data Pipeline &amp; Preprocessing","From raw sensor data to 18-dimensional model-ready feature vector")}
  <div style="flex:1;display:flex;flex-direction:column;gap:10px;">
    <div style="flex:1;">{img("dp","full")}</div>
    <div class="grid5">{step_cards}</div>
  </div>
""", 5))

# ── S6: Input Features Table ──────────────────────────────────────────────────
slides.append(slide(f"""
  {hdr("Input Features","8 numeric features + 1 categorical feature (crop type)")}
  {table_html(
    ["Feature","Type","Unit","Range","Description"],
    [["Nitrogen (N)","Numeric","kg/ha","0–200","Available nitrogen in soil"],
     ["Phosphorus (P)","Numeric","kg/ha","0–150","Available phosphorus in soil"],
     ["Potassium (K)","Numeric","kg/ha","0–120","Available potassium in soil"],
     ["pH","Numeric","—","3.0–10.0","Soil acidity / alkalinity"],
     ["Moisture","Numeric","%","5–100","Soil moisture percentage"],
     ["Temperature","Numeric","°C","-5–50","Ambient temperature"],
     ["Rainfall","Numeric","mm","0–450","Monthly rainfall"],
     ["Humidity","Numeric","%","10–100","Air humidity percent"],
     ["Crop Type","Categorical","—","10 types","Rice, Wheat, Maize, Cotton, …"]]
  )}
""", 6))

# ── S7: Crop Profiles ─────────────────────────────────────────────────────────
slides.append(slide(f"""
  {hdr("Supported Crop Profiles","Optimal NPK, pH, temperature and yield ranges for each crop")}
  {table_html(
    ["Crop","N (kg/ha)","P (kg/ha)","K (kg/ha)","pH","Temp (°C)","Yield (t/ha)"],
    [["Rice","60–120","30–60","30–60","5.5–7.0","22–35","3.0–8.0"],
     ["Wheat","80–140","40–70","20–50","6.0–7.5","12–25","2.5–6.5"],
     ["Maize","80–160","30–60","20–50","5.8–7.0","18–32","3.0–9.0"],
     ["Cotton","60–120","20–50","20–40","6.0–7.5","20–35","1.5–4.0"],
     ["Sugarcane","100–200","40–80","40–80","5.5–7.5","22–36","50–120"],
     ["Soybean","20–50","40–70","20–50","6.0–7.0","20–30","1.5–4.0"],
     ["Potato","80–150","50–90","60–100","5.0–6.5","15–25","15–40"],
     ["Tomato","60–130","40–80","50–90","5.5–7.0","18–30","20–60"],
     ["Groundnut","10–30","30–60","20–40","5.5–7.0","22–33","1.0–3.5"],
     ["Barley","60–110","30–55","20–45","6.0–8.0","10–22","2.0–5.5"]]
  )}
""", 7))

# ── S8: Random Forest ────────────────────────────────────────────────────────
slides.append(slide(f"""
  {hdr("Model 1: Random Forest Classifier","Fertilizer type recommendation — ensemble of 15 decision trees","#38BDF8")}
  <div class="row">
    <div class="col card b" style="padding:12px;max-width:42%;">
      <div class="sct" style="color:#38BDF8;">How It Works</div>
      {bullets(["Bootstrap Sampling — each tree gets a random sample (with replacement)",
               "Random Feature Subset — 60% of features per split for diversity",
               "Gini Impurity — measures split quality at each node",
               "Majority Vote — aggregates predictions from all 15 trees",
               "Confidence Score — % of trees agreeing on the top class",
               "7 Output Classes — one per fertilizer type"])}
      <div class="sct" style="color:#00C97B;margin-top:10px;">Key Params</div>
      {table_html(["Parameter","Value"],[["Trees","15"],["Max Depth","12"],["Min Samples/Leaf","3"],["Feature Ratio","60%"],["Criterion","Gini"]])}
    </div>
    <div class="col">{img("rf")}</div>
  </div>
""", 8))

# ── S9: Gradient Boosting ─────────────────────────────────────────────────────
slides.append(slide(f"""
  {hdr("Model 2: Gradient Boosting Regressor","Yield prediction using 50 sequential decision stumps","#F5A623")}
  <div class="row">
    <div class="col card o" style="padding:12px;max-width:42%;">
      <div class="sct" style="color:#F5A623;">Algorithm Steps</div>
      {bullets(["Initialize with mean of all target yields",
               "For each of 50 estimators:",
               "  a) Compute residuals (actual – prediction)",
               "  b) Fit a decision stump (depth 4) on residuals",
               "  c) Predictions += learning_rate × stump output",
               "Confidence = 1 − (RMSE / target std deviation)",
               "Output: yield (t/ha), confidence, RMSE"])}
      <div class="sct" style="color:#00C97B;margin-top:10px;">Key Params</div>
      {table_html(["Parameter","Value"],[["Estimators","50"],["Learning Rate","0.1"],["Max Depth","4"],["Min Samples/Split","5"],["Loss","MSE"]])}
    </div>
    <div class="col">{img("gb")}</div>
  </div>
""", 9))

# ── S10: Fertilizer Types ─────────────────────────────────────────────────────
slides.append(slide(f"""
  {hdr("Fertilizer Types &amp; Nutrient Content","7 types with NPK composition and real market cost data")}
  {table_html(
    ["Fertilizer","N (%)","P (%)","K (%)","Cost (₹/kg)","Best For"],
    [["Urea","46","0","0","8","High N deficiency"],
     ["DAP","18","46","0","28","N + P deficiency"],
     ["NPK 10-26-26","10","26","26","22","P + K deficiency"],
     ["NPK 20-20-20","20","20","20","25","Balanced needs"],
     ["MOP","0","0","60","18","High K deficiency"],
     ["SSP","0","16","0","10","Phosphorus deficiency"],
     ["Amm. Sulphate","21","0","0","12","Acidic soils"]]
  )}
""", 10))

# ── S11: Selection Logic ──────────────────────────────────────────────────────
node_style = "border-radius:8px;padding:8px 12px;text-align:center;font-size:8pt;font-weight:600;"
slides.append(slide(f"""
  {hdr("Fertilizer Selection Logic","Rule-based training label assignment — decision tree flowchart")}
  <div style="flex:1;display:flex;flex-direction:column;align-items:center;justify-content:space-around;gap:6px;padding-top:4px;">
    <div style="{node_style}background:#1e3a5f;border:1.5px solid #38BDF8;color:#38BDF8;width:280px;">N &lt; 50 AND P &lt; 35 AND K &lt; 30?</div>
    <div style="display:flex;gap:100px;width:100%;justify-content:center;">
      <div style="{node_style}background:#1A3C2A;border:1.5px solid #00C97B;color:#00C97B;width:180px;">✅ YES → NPK 20-20-20<br><span style="color:#8B9CB2;font-weight:400;">(Balanced)</span></div>
      <div style="{node_style}background:#1e3a5f;border:1.5px solid #38BDF8;color:#38BDF8;width:180px;">NO → N &lt; 50 AND P &lt; 35?</div>
    </div>
    <div style="display:flex;gap:70px;width:100%;justify-content:center;">
      <div style="{node_style}background:#1A3C2A;border:1.5px solid #00C97B;color:#00C97B;width:150px;">✅ YES → DAP (N+P)</div>
      <div style="{node_style}background:#1e3a5f;border:1.5px solid #38BDF8;color:#38BDF8;width:150px;">NO → Only N &lt; 50?</div>
    </div>
    <div style="display:flex;gap:50px;width:100%;justify-content:center;">
      <div style="{node_style}background:#1A3C2A;border:1.5px solid #00C97B;color:#00C97B;width:140px;">✅ YES → Urea (N only)</div>
      <div style="{node_style}background:#3a2a10;border:1.5px solid #F5A623;color:#F5A623;width:260px;font-size:7.5pt;">NO → K&lt;30→MOP | P&lt;35→SSP | pH&lt;5.5→Am.Sul. | else→NPK 10-26-26</div>
    </div>
  </div>
  <div style="text-align:center;font-size:7.5pt;color:#8B9CB2;margin-top:4px;">🔵 Decision node (rule check)   ✅ Leaf node (fertilizer assigned)   🟠 Multiple outcomes</div>
""", 11))

# ── S12: Optimization Layer ───────────────────────────────────────────────────
opt_cols = ""
for icon, title, items, cls, clr in [
    ("💊","Quantity Optimization",["Compute N,P,K deficits vs crop profile","Qty = max deficit / nutrient% × 100","Capped at 20–500 kg/hectare","Scales by user's land area"],"g","#00C97B"),
    ("💰","Cost Estimation",["Cost = quantity × cost-per-kg","Per-hectare and total cost","7 fertilizers with real market rates","Supports ₹ INR currency"],"o","#F5A623"),
    ("🌍","Environmental Impact",["Over-application penalty (+15)","High rainfall leaching risk (+20)","Soil acidification risk (+15)","Moisture runoff risk (+10)","Qty-proportional impact (+0–40)"],"p","#E84D8A"),
]:
    opt_cols += f'<div class="card {cls}" style="padding:12px;"><div class="sct" style="color:{clr};margin-bottom:6px;">{icon} {title}</div>{bullets(items)}</div>'

slides.append(slide(f"""
  {hdr("Optimization Layer","Calculating optimal quantity, cost, and environmental impact score")}
  <div class="grid3">{opt_cols}</div>
""", 12))

# ── S13: Climate Engine ───────────────────────────────────────────────────────
slides.append(slide(f"""
  {hdr("Climate-Aware Decision Engine","7-day weather forecast adjusts quantity and application timing","#38BDF8")}
  <div class="row">
    <div class="col card b" style="padding:12px;max-width:42%;">
      <div class="sct" style="color:#38BDF8;">🌦️ Climate Logic</div>
      {bullets(["Generates 7-day forecast: temp, rainfall, humidity, wind",
               "Analyzes next 5 days for best application window",
               "Heavy rain (&gt;50mm) → reduce quantity 15%, delay",
               "Moderate rain (&gt;25mm) → reduce 8%, split doses",
               "High temp (&gt;35°C) → apply early morning only",
               "High humidity (&gt;80%) → monitor fungal disease risk",
               "Best Application Day — lowest risk window is scored"])}
      <div style="margin-top:8px;">
      {table_html(["Condition","Qty Adjust","Risk"],[
        ["Rain &gt;50mm","−15%","High"],["Rain 25–50mm","−8%","Medium"],
        ["Temp &gt;35°C","−10%","Medium"],["Humidity &gt;80%","No change","Low"],
        ["Dry + Moderate","Normal","Low"]])}
      </div>
    </div>
    <div class="col">{img("ce")}</div>
  </div>
""", 13))

# ── S14: Risk Prediction ──────────────────────────────────────────────────────
slides.append(slide(f"""
  {hdr("Risk Prediction Module","7 risk categories with severity, probability, and actionable recommendations","#E84D8A")}
  <div class="row">
    <div class="col" style="max-width:50%;">
      {table_html(
        ["Risk Category","Trigger Condition","Severity","Recommendation"],
        [["⚠️ Over-Fertilization","NPK excess &gt;20 kg/ha","High/Med","Reduce, soil test"],
         ["⚖️ Nutrient Imbalance","N:P:K ratio gap &gt;0.6","High/Med","Apply balanced NPK"],
         ["📉 Low Yield","Predicted &lt;60% of avg","High/Med","Review soil + irrigation"],
         ["🧪 pH Stress","pH outside optimal ±1","High/Med","Lime or Sulfur"],
         ["🏜️ Drought Stress","Moisture &lt;25%, Rain &lt;30","High/Med","Irrigate + mulch"],
         ["🌊 Leaching","Rain &gt;200mm, Qty &gt;100","High/Med","Slow-release doses"],
         ["🌡️ Temp Stress","Temp outside ±5°C","Medium","Shade / row covers"]]
      )}
    </div>
    <div class="col">{img("rr")}</div>
  </div>
""", 14))

# ── S15: Sustainability Index ─────────────────────────────────────────────────
sus_cards = ""
for icon, name, wt, formula in [
    ("🌿","Fertilizer Efficiency","20%","100 − qty/5"),
    ("⚖️","Soil Balance","15%","100 − total deficit"),
    ("🌍","Environmental","20%","100 − env impact"),
    ("💧","Water Impact","15%","Rain + moisture penalties"),
    ("🌾","Soil Health Trend","15%","From degradation tracker"),
    ("♻️","Carbon Footprint","15%","100 − (qty × 0.5)/2.5"),
]:
    sus_cards += (f'<div class="card g" style="text-align:center;padding:10px;display:flex;flex-direction:column;align-items:center;justify-content:center;">'
                  f'<div style="font-size:20pt;">{icon}</div>'
                  f'<div style="font-size:8.5pt;font-weight:700;color:#00C97B;margin:4px 0;">{name}</div>'
                  f'<div style="font-size:7.5pt;color:#F5A623;">Weight: {wt}</div>'
                  f'<div style="font-size:7pt;color:#CBD5E1;margin-top:3px;">{formula}</div>'
                  f'</div>')
slides.append(slide(f"""
  {hdr("Sustainability Index","6-component weighted scoring system with grades A–F")}
  <div class="grid6">{sus_cards}</div>
  <div style="text-align:center;font-size:7.5pt;color:#8B9CB2;margin-top:8px;">
    Grades: &nbsp; A (≥80) Excellent &nbsp;|&nbsp; B (≥65) Good &nbsp;|&nbsp; C (≥50) Fair &nbsp;|&nbsp; D (≥35) Poor &nbsp;|&nbsp; F (&lt;35) Critical
  </div>
""", 15))

# ── S16: Soil Health Tracker ──────────────────────────────────────────────────
slides.append(slide(f"""
  {hdr("Soil Health Degradation Tracker","Linear trend projection — predicts nutrient decline over 6+ months")}
  <div class="row">
    <div class="col card g" style="padding:12px;max-width:42%;">
      <div class="sct" style="color:#00C97B;">📈 How It Works</div>
      {bullets(["Records soil N, P, K, pH, moisture for each analysis",
               "Builds time-series history (up to 100 entries) in localStorage",
               "Linear regression on nutrient values over time index",
               "Projects values 6+ months into the future",
               "Classifies trend: Declining / Stable / Improving",
               "Health score (0–100) from nutrient proximity to optimal",
               "Generates urgency-based recommendations"])}
      <div style="margin-top:8px;">
      {table_html(["Nutrient","Optimal Low","Optimal High"],
        [["Nitrogen","40 kg/ha","120 kg/ha"],["Phosphorus","30 kg/ha","70 kg/ha"],
         ["Potassium","25 kg/ha","60 kg/ha"],["pH","5.5","7.5"]])}
      </div>
    </div>
    <div class="col">{img("sh")}</div>
  </div>
""", 16))

# ── S17: Adaptive Learning ────────────────────────────────────────────────────
slides.append(slide(f"""
  {hdr("Adaptive Learning &amp; Feedback Loop","Continuous self-improvement through farmer feedback","#A78BFA")}
  <div class="row">
    <div class="col card pu" style="padding:12px;max-width:42%;">
      <div class="sct" style="color:#A78BFA;">🧠 Adaptive Learning</div>
      {bullets(["Records every recommendation (input + predicted output)",
               "Stores up to 200 recommendations in localStorage",
               "Correction factor = actual_yield / predicted_yield (per crop)",
               "Blends: 70% model + 30% learned correction",
               "Blend weight grows with more feedback (max 30%)",
               "Accuracy tracked: within 20% error = 'correct'",
               "Builds personal crop profiles over time"])}
      <div class="sct" style="color:#00C97B;margin-top:10px;">📝 Farmer Feedback</div>
      {bullets(["Submit: actual yield, fertilizer used, satisfaction (1–5)",
               "Status: Awaiting → Learning (1+) → Active (5+ feedbacks)",
               "Trend: first-half vs second-half accuracy tracking"])}
    </div>
    <div class="col">{img("al")}</div>
  </div>
""", 17))

# ── S18: Scenario Simulator ───────────────────────────────────────────────────
slides.append(slide(f"""
  {hdr("Scenario Simulator — What-If Analysis","Explore impact of different conditions on yield, cost, and sustainability")}
  {table_html(
    ["Scenario","Modifications Applied","Use Case"],
    [["📉 Reduce Fertilizer 30%","N, P, K × 0.7","Cost saving analysis"],
     ["📈 Increase Fertilizer 30%","N, P, K × 1.3","Yield maximization"],
     ["🏜️ Drought Conditions","Rainfall=20mm, Moisture=20%","Drought preparedness"],
     ["🌧️ Heavy Rainfall","Rainfall=350mm, Moisture=85%","Flood risk planning"],
     ["🌡️ Heat Wave","Temp=42°C, Humidity=30%","Extreme heat response"],
     ["✨ Optimal Conditions","All params set to ideal values","Best-case benchmarking"]]
  )}
  <div style="font-size:7.5pt;color:#8B9CB2;margin-top:8px;">
    Each scenario computes: Δ Yield (t/ha) &nbsp;|&nbsp; Δ Cost (₹) &nbsp;|&nbsp; Δ Environmental Impact &nbsp;|&nbsp; Natural-language recommendation
  </div>
""", 18))

# ── S19: Explainability & Smart Alerts ───────────────────────────────────────
slides.append(slide(f"""
  {hdr("Explainability &amp; Smart Alerts","Every recommendation comes with human-readable reasoning and risk alerts")}
  <div class="row">
    <div class="col card b" style="padding:12px;">
      <div class="sct" style="color:#38BDF8;">💡 Model Explainability</div>
      {bullets(["Feature importance via correlation analysis","Top-3 influencing features shown per recommendation",
               "Nutrient-specific explanations (N, P, K deficit context)","Environmental context notes",
               "Yield context: above/below average for that crop type","Sustainability badges earned"])}
    </div>
    <div class="col card o" style="padding:12px;">
      <div class="sct" style="color:#F5A623;">🔔 Smart Alerts</div>
      {table_html(
        ["Level","Trigger","Message"],
        [["🔴 Critical","pH &lt; 4.5","Apply lime immediately"],
         ["🔴 Critical","Moisture &lt; 20%","Irrigate before applying"],
         ["🟡 Warning","pH &lt; 5.5","Acidic soil detected"],
         ["🟡 Warning","pH &gt; 8.5","Apply gypsum to correct"],
         ["🟡 Warning","N &lt; 20 kg/ha","Yellowing leaves risk"],
         ["🟡 Warning","Temp &gt; 40°C","Apply in early morning"],
         ["🟢 Info","Low deficits","Minimal fertilizer needed"]]
      )}
    </div>
  </div>
""", 19))

# ── S20: Tech Stack & Key Metrics ─────────────────────────────────────────────
slides.append(slide(f"""
  {hdr("Technology Stack &amp; Key Metrics","Engineering details and performance characteristics")}
  <div class="row">
    <div class="col card b" style="padding:12px;">
      <div class="sct" style="color:#38BDF8;">⚡ Technology Stack</div>
      {bullets(["Frontend: Vanilla JavaScript (ES6 Modules)","Styling: CSS3 with Glassmorphism design",
               "Build: Vite bundler and dev server","ML: Custom Random Forest + Gradient Boosting (no libraries)",
               "Storage: localStorage for adaptive learning + history","Charts: Custom SVG / Canvas charting",
               "Deployment: Static hosting (Vercel / Netlify)","Zero server dependencies — 100% in browser"])}
    </div>
    <div class="col card g" style="padding:12px;">
      <div class="sct" style="color:#00C97B;">📊 Key Metrics</div>
      {table_html(
        ["Metric","Value"],
        [["Training Samples","2,000 synthetic"],["Feature Dimensions","18 (8 numeric + 10 encoded)"],
         ["Supported Crops","10"],["Fertilizer Classes","7"],["RF Trees","15"],
         ["GB Estimators","50"],["Risk Categories","7"],["Sustainability Components","6"],
         ["Code Modules","12"]]
      )}
    </div>
  </div>
""", 20))

# ── S21: Roadmap ──────────────────────────────────────────────────────────────
phase_cards = ""
for phase, title, desc, cls, clr in [
    ("Phase 1","Real Weather API","Connect to IMD / OpenWeather for live 7-day forecast data","g","#00C97B"),
    ("Phase 2","Satellite Imagery","NDVI from Sentinel-2 for real-time crop health monitoring","b","#38BDF8"),
    ("Phase 3","Backend + Auth","Node.js / Firebase for cloud data persistence and user accounts","o","#F5A623"),
    ("Phase 4","Mobile App","Offline-capable React Native app for rural farmers with SMS alerts","p","#E84D8A"),
]:
    phase_cards += (f'<div class="card {cls}" style="text-align:center;padding:14px;display:flex;flex-direction:column;align-items:center;justify-content:center;">'
                    f'<div style="font-size:8.5pt;font-weight:700;color:{clr};">{phase}</div>'
                    f'<div style="font-size:12pt;font-weight:700;margin:6px 0;">{title}</div>'
                    f'<div style="font-size:8pt;color:#CBD5E1;">{desc}</div>'
                    f'</div>')
slides.append(slide(f"""
  {hdr("Future Roadmap","Planned enhancements for the next development phases")}
  <div class="grid4">{phase_cards}</div>
""", 21))

# ── S22: Thank You ────────────────────────────────────────────────────────────
slides.append(slide(f"""
  <div class="deco" style="width:400px;height:400px;background:#00C97B;top:50%;left:50%;transform:translate(-50%,-50%);opacity:.06;"></div>
  <div style="flex:1;display:flex;flex-direction:column;align-items:center;justify-content:center;text-align:center;">
    <div style="font-size:36pt;font-weight:900;color:#00C97B;">Thank You!</div>
    <div style="font-size:14pt;font-weight:600;margin-top:14px;">AgriML — Empowering Farmers with AI</div>
    <div style="font-size:10pt;color:#8B9CB2;margin-top:12px;">Questions &amp; Discussion</div>
    <div style="font-size:9pt;color:#38BDF8;margin-top:24px;">
      🌾 Smart Fertilizer &nbsp;•&nbsp; 📊 Yield Prediction &nbsp;•&nbsp; 🌦️ Climate Aware &nbsp;•&nbsp; ♻️ Sustainable
    </div>
  </div>
""", 22))

# ── Assemble HTML ─────────────────────────────────────────────────────────────
html = f"""<!DOCTYPE html>
<html lang="en">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width,initial-scale=1.0">
<title>AgriML — AI-Powered Fertilizer Recommendation System</title>
<style>{CSS}</style>
</head>
<body>
{''.join(slides)}
</body>
</html>"""

# Write HTML file
with open(OUT_HTML, "w", encoding="utf-8") as f:
    f.write(html)
print(f"[OK] HTML written -> {OUT_HTML}")

# ── Chrome headless → PDF ─────────────────────────────────────────────────────
chrome = find_chrome()
if not chrome:
    print("[ERROR] Chrome not found! Please install Google Chrome.")
    sys.exit(1)

print(f"[INFO] Using Chrome: {chrome}")
print(f"[INFO] Generating PDF -> {OUT_PDF}")

html_url = "file:///" + OUT_HTML.replace("\\", "/")

result = subprocess.run([
    chrome,
    "--headless=new",
    "--disable-gpu",
    "--no-sandbox",
    "--disable-web-security",
    f"--print-to-pdf={OUT_PDF}",
    "--no-pdf-header-footer",
    "--run-all-compositor-stages-before-draw",
    "--virtual-time-budget=5000",
    html_url
], capture_output=True, text=True, timeout=60)

if not os.path.exists(OUT_PDF):
    # Fallback: try older --headless flag
    print("[INFO] Retrying with legacy --headless flag...")
    result = subprocess.run([
        chrome,
        "--headless",
        "--disable-gpu",
        "--no-sandbox",
        f"--print-to-pdf={OUT_PDF}",
        "--print-to-pdf-no-header",
        html_url
    ], capture_output=True, text=True, timeout=60)

if os.path.exists(OUT_PDF):
    size_kb = os.path.getsize(OUT_PDF) // 1024
    print(f"[OK] PDF generated ({size_kb} KB) -> {OUT_PDF}")
    # Open in Chrome
    subprocess.Popen([chrome, OUT_PDF])
    print("[OK] Opened in Chrome!")
else:
    print("[ERROR] PDF generation failed.")
    print("STDOUT:", result.stdout[:500])
    print("STDERR:", result.stderr[:500])
    # Fallback: open HTML in Chrome for manual print
    print("[INFO] Opening HTML in Chrome — use Ctrl+P → Save as PDF instead.")
    subprocess.Popen([chrome, html_url])
