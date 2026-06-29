"""
Azure Pricing Calculator → Cost Estimation Workbook
━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
• Parses any Azure Pricing Calculator export (.xlsx)
• Smart API Fetching: Includes intelligent SKU fallbacks for Azure API quirks.
• Perfect Mathematical Deduction: Forces Compute, OS, and SQL components 
  to exactly equal the Calculator's total row, preventing any cost inflation.
• Standalone License Bypassing: Detects detached RHEL/SUSE/SQL rows and 
  prints them cleanly as single items.
• Safe parsing for missing service types and missing files.

Usage:
    python convert.py input.xlsx [output.xlsx]
"""

import re, sys, time, logging
from pathlib import Path

import requests
from openpyxl import load_workbook, Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side

# ── Logging ──────────────────────────────────────────────────────────────────
logging.basicConfig(
    level=logging.INFO,
    format="  %(message)s",
    handlers=[logging.StreamHandler(sys.stdout)]
)
log = logging.getLogger(__name__)

# ── Style helpers ─────────────────────────────────────────────────────────────
THIN   = Side(style="thin", color="000000")
BORDER = Border(left=THIN, right=THIN, top=THIN, bottom=THIN)
NUM    = "#,##0.00"

def _f(bold=False, italic=False, size=11, color="000000"):
    return Font(name="Calibri", bold=bold, italic=italic, size=size, color=color)

def _fill(h): return PatternFill("solid", fgColor=h)
def _al(h="left", v="center", w=False): return Alignment(horizontal=h, vertical=v, wrap_text=w)

def hdr(c, v, wrap=False):
    c.value=v; c.font=Font(name="Calibri",bold=True,size=10,color="FFFFFF")
    c.fill=_fill("4472C4"); c.alignment=_al("center","center",wrap); c.border=BORDER

def dat(c, v, bold=False, italic=False, align="left", color="000000"):
    c.value=v; c.font=_f(bold,italic,color=color)
    c.alignment=_al(align,"center",True); c.border=BORDER
    if isinstance(v,(int,float)) and align=="right": c.number_format=NUM

def tot(c, v):
    c.value=v; c.font=_f(bold=True); c.fill=_fill("D9E1F2")
    c.alignment=_al("right","center"); c.border=BORDER
    if isinstance(v,(int,float)): c.number_format=NUM

def widths(ws, d):
    for col,w in d.items(): ws.column_dimensions[col].width=w

# ── Resource → sheet name ────────────────────────────────────────────────────
SVC_MAP = {
    "virtual machines":          "Virtual Machines",
    "managed disks":             "Managed Disks",
    "azure backup":              "Azure Backup",
    "backup":                    "Azure Backup",
    "load balancer":             "Load Balancer",
    "load balancers":            "Load Balancer",
    "application gateway":       "Application Gateway",
    "azure firewall":            "Azure Firewall",
    "vpn gateway":               "VPN Gateway",
    "storage":                   "Storage Accounts",
    "storage accounts":          "Storage Accounts",
    "azure site recovery":       "Azure Site Recovery",
    "azure virtual desktop":     "Azure Virtual Desktop",
    "sql database":              "SQL",
    "azure sql":                 "SQL",
    "sql":                       "SQL",
    "ip addresses":              "Public IP Addresses",
    "public ip addresses":       "Public IP Addresses",
    "bandwidth":                 "Bandwidth",
    "azure dns":                 "Azure DNS",
    "azure monitor":             "Azure Monitor",
    "key vault":                 "Key Vault",
}

SHEET_ORDER = [
    "Virtual Machines","Managed Disks","Public IP Addresses",
    "Load Balancer","Application Gateway","Azure Firewall","VPN Gateway",
    "Storage Accounts","Azure Backup","Azure Site Recovery",
    "Azure Virtual Desktop","SQL","Bandwidth","Azure DNS",
    "Azure Monitor","Key Vault","Others",
]

SKIP = {"support","disclaimer","total","licensing program",
        "billing account","billing profile"}

# ── Azure region display → ARM name ─────────────────────────────────────────
REGION_MAP = {
    "central india":"centralindia","east us":"eastus","east us 2":"eastus2",
    "west us":"westus","north europe":"northeurope","west europe":"westeurope",
    "southeast asia":"southeastasia","uk south":"uksouth",
    "australia east":"australiaeast","japan east":"japaneast",
    "canada central":"canadacentral","south india":"southindia",
}

def arm_region(display):
    return REGION_MAP.get(display.lower().strip(), display.lower().strip().replace(" ",""))

# ══════════════════════════════════════════════════════════════════════════════
#  AZURE RETAIL PRICING API
# ══════════════════════════════════════════════════════════════════════════════
API = "https://prices.azure.com/api/retail/prices"
_cache = {}

def _api(filt, currency="INR"):
    key = filt+currency
    if key in _cache: return _cache[key]
    try:
        r = requests.get(API, params={"api-version":"2023-01-01-preview","$filter":filt,"currencyCode":currency}, timeout=20)
        r.raise_for_status()
        items = r.json().get("Items",[])
        _cache[key] = items
        return items
    except Exception as e:
        log.warning(f"API error: {e}")
        return []

def _hourly_to_monthly(price): return price * 730

def _pick_linux(items):
    cands = [i for i in items if "windows" not in i.get("productName", "").lower() and "spot" not in i.get("meterName", "").lower() and "low priority" not in i.get("meterName", "").lower()]
    prices = [i["retailPrice"] for i in cands if i.get("retailPrice", 0) > 0]
    return min(prices) if prices else None

def _pick_windows(items):
    cands = [i for i in items if "windows" in i.get("productName", "").lower() and "spot" not in i.get("meterName", "").lower() and "low priority" not in i.get("meterName", "").lower()]
    prices = [i["retailPrice"] for i in cands if i.get("retailPrice", 0) > 0]
    return min(prices) if prices else None

def _pick_ri(items):
    cands = [i for i in items if "spot" not in i.get("meterName", "").lower()]
    prices = [i["retailPrice"] for i in cands if i.get("retailPrice", 0) > 0]
    return min(prices) if prices else None

def get_vm_pricing(sku, region_display, currency="INR"):
    region = arm_region(region_display)
    result = {"compute_payg": None, "windows_license": None, "compute_ri1": None, "compute_ri3": None}

    try:
        def fetch_payg(s): return _api(f"armSkuName eq '{s}' and armRegionName eq '{region}' and priceType eq 'Consumption'", currency)
        def fetch_ri(s):   return _api(f"armSkuName eq '{s}' and armRegionName eq '{region}' and priceType eq 'Reservation'", currency)

        # Intelligent SKU Fallback (Azure API sometimes drops the 's' for standard/premium storage)
        payg_items = fetch_payg(sku)
        if not payg_items: payg_items = fetch_payg(sku.replace("s_v", "_v").replace("s_V", "_V"))
        if not payg_items: payg_items = fetch_payg(sku.replace("ds_v", "d_v").replace("ds_V", "d_V"))
        if not payg_items: payg_items = fetch_payg(sku.replace("ds_v", "s_v").replace("ds_V", "s_V"))

        ri_items = fetch_ri(sku)
        if not ri_items: ri_items = fetch_ri(sku.replace("s_v", "_v").replace("s_V", "_V"))
        if not ri_items: ri_items = fetch_ri(sku.replace("ds_v", "d_v").replace("ds_V", "d_V"))
        if not ri_items: ri_items = fetch_ri(sku.replace("ds_v", "s_v").replace("ds_V", "s_V"))

        linux_hr = _pick_linux(payg_items)
        win_hr   = _pick_windows(payg_items)

        api_comp = _hourly_to_monthly(linux_hr) if linux_hr else None
        api_win_tot = _hourly_to_monthly(win_hr) if win_hr else None
        
        result["compute_payg"] = api_comp
        if api_win_tot and api_comp:
            result["windows_license"] = max(0, api_win_tot - api_comp)

        ri1_cands = [i for i in ri_items if i.get("reservationTerm") == "1 Year"]
        ri1_val = _pick_ri(ri1_cands)
        if ri1_val is not None and api_comp is not None:
            if ri1_val > (api_comp * 3): result["compute_ri1"] = ri1_val / 12
            else: result["compute_ri1"] = ri1_val

        ri3_cands = [i for i in ri_items if i.get("reservationTerm") == "3 Years"]
        ri3_val = _pick_ri(ri3_cands)
        if ri3_val is not None and api_comp is not None:
            if ri3_val > (api_comp * 10): result["compute_ri3"] = ri3_val / 36
            else: result["compute_ri3"] = ri3_val

    except Exception as e:
        log.warning(f"    VM pricing error for {sku}: {e}")

    return result

def extract_vm_sku(desc):
    m = re.match(r'^\s*[\d,]+\s+([^()]+)\(', desc)
    if m:
        raw = m.group(1).strip()
        norm = re.sub(r'[\s\-]+', '_', raw)
        if not norm.lower().startswith("standard_"):
            norm = "Standard_" + norm
        return norm
        
    patterns = [r'^\d+\s+((?:[A-Z][A-Za-z0-9]+\s+)+v\d+)', r'^\d+\s+([A-Z][A-Za-z0-9]+)\s*\(', r'(Standard_[A-Za-z0-9_]+)']
    for pat in patterns:
        m = re.search(pat, desc.strip())
        if m:
            raw = m.group(1).strip()
            norm = re.sub(r'\s+', '_', raw)
            if not norm.startswith("Standard_"): norm = "Standard_" + norm
            return norm
    return None

def extract_quantity(desc):
    m = re.match(r'^\s*([0-9,]+)\s+', desc)
    if m:
        try:
            q = int(m.group(1).replace(',', ''))
            return q if q > 0 else 1
        except: return 1
    return 1

def detect_os(desc):
    desc_l = desc.lower()
    if "hybrid benefit" in desc_l: return "Linux" 
    if "red hat" in desc_l or "rhel" in desc_l: return "Red Hat"
    if "suse" in desc_l or "sles" in desc_l: return "SUSE"
    if "linux" in desc_l or "ubuntu" in desc_l or "centos" in desc_l: return "Linux"
    if "windows" in desc_l: return "Windows"
    return "Windows" 

def detect_sql_license(desc):
    desc_l = desc.lower()
    if "sql enterprise" in desc_l: return "SQL Enterprise License"
    if "sql standard"   in desc_l: return "SQL Standard License"
    if "sql web"        in desc_l: return "SQL Web License"
    if "sql developer"  in desc_l: return "SQL Developer License"
    if "sql" in desc_l: return "SQL License"
    return None

def get_exact_license_name(desc, os_type):
    parts = re.split(r'[,;]', desc)
    sql_name = os_name = None
    for p in parts:
        p_lower = p.lower()
        if "sql" in p_lower:
            sql_name = re.sub(r'\s*\([^)]*\)', '', p).strip() 
        if os_type in ["Red Hat", "SUSE"] and ("red hat" in p_lower or "rhel" in p_lower or "suse" in p_lower or "sles" in p_lower):
            os_name = re.sub(r'\s*\([^)]*\)', '', p).strip()
    return sql_name, os_name

# ══════════════════════════════════════════════════════════════════════════════
#  INPUT PARSING
# ══════════════════════════════════════════════════════════════════════════════
def parse_format_a(wb):
    ws = wb.active
    rows = []
    in_data = False
    for r in ws.iter_rows(values_only=True):
        v = list(r)
        if v[0] == "Service category": in_data=True; continue
        if not in_data: continue
        svc_cat, svc_type, region, desc = str(v[0] or "").strip(), str(v[1] or "").strip(), str(v[3] or "").strip(), str(v[4] or "").strip()
        cost_raw = v[5]
        if svc_cat.lower() in SKIP or region.lower() in SKIP or desc.lower() in SKIP or not isinstance(cost_raw,(int,float)) or (not svc_cat and not svc_type): continue
        rows.append({
            "svc_cat": svc_cat, "svc_type": svc_type, "cust_name": str(v[2] or "").strip(),
            "region": region, "desc": desc, "payg": float(cost_raw),
            "ri1": None, "ri3": None, "remarks": "", "sub_rows": [], "fmt": "A"
        })
    return rows

def parse_format_b(wb):
    rows = []
    for sname in wb.sheetnames:
        if sname.lower() == "summary": continue
        ws = wb[sname]
        in_data = False
        for r in ws.iter_rows(values_only=True):
            v = list(r)
            if str(v[0] or "").lower() == "service category": in_data=True; continue
            if not in_data: continue
            desc = str(v[4] or "").strip()
            if not desc or desc.lower() == "total": continue
            svc_cat, svc_type, cust, region = str(v[0] or "").strip(), str(v[1] or "").strip(), str(v[2] or "").strip(), str(v[3] or "").strip()
            remarks = str(v[8] or "").strip() if len(v)>8 else ""
            
            def flt(x): return float(x) if x is not None else None
            payg, ri1, ri3 = flt(v[5]), flt(v[6]), flt(v[7])
            
            if (not svc_cat and not svc_type and not cust and not region):
                if rows and payg is not None:
                    rows[-1]["sub_rows"].append({"desc":desc, "payg":payg, "ri1":ri1 or payg, "ri3":ri3 or payg, "remarks":remarks})
            else:
                if payg is None: continue
                rows.append({
                    "svc_cat":svc_cat, "svc_type":svc_type, "cust_name":cust, "region":region, "desc":desc,
                    "payg":payg, "ri1":ri1 or payg, "ri3":ri3 or payg, "remarks":remarks, "sub_rows":[], "fmt":"B"
                })
    return rows

def classify(rows):
    buckets = {}
    for r in rows:
        # Safely convert to lower case even if the cell is blank/None
        key = str(r.get("svc_type") or "").lower().strip()
        sheet = SVC_MAP.get(key)
        if not sheet:
            for k,v in SVC_MAP.items():
                if k in key: sheet=v; break
        buckets.setdefault(sheet or "Others", []).append(r)
    return buckets

# ══════════════════════════════════════════════════════════════════════════════
#  API ENRICHMENT & BULLETPROOF MATH DEDUCTION
# ══════════════════════════════════════════════════════════════════════════════
def enrich_vms(vm_rows, currency="INR"):
    log.info(f"Querying Azure Retail Pricing API for {len(vm_rows)} VM(s)...")
    for row in vm_rows:
        desc, region = row["desc"], row["region"]
        os_type, sql_lbl = detect_os(desc), detect_sql_license(desc)
        qty = extract_quantity(desc)
        sku = extract_vm_sku(desc)

        sql_exact, os_exact = get_exact_license_name(desc, os_type)
        row["sql_lbl_exact"] = sql_exact or "SQL License"
        row["os_lbl_exact"]  = os_exact or f"{os_type} License"
        
        row["api"] = {}

        if not sku:
            is_standalone = False
            p = {}
            if os_type in ["Red Hat", "SUSE"] or sql_lbl:
                is_standalone = True

            if is_standalone:
                p["is_standalone"] = True
                log.info(f"    Standalone License Detected | {row['os_lbl_exact']} / {row['sql_lbl_exact']}")
                row["api"] = p
            else:
                log.warning(f"    Could not extract SKU from: {desc[:60]}")
            continue

        sqllog = f" | {sql_exact}" if sql_exact else ""
        log.info(f"    {sku} (Qty: {qty}) | {region} | {os_type}{sqllog}")
        
        p = get_vm_pricing(sku, region, currency)
        
        api_comp = (p.get("compute_payg") or 0) * qty
        api_win  = (p.get("windows_license") or 0) * qty
        
        orig_payg = row.get("payg", 0)
        compute_payg = api_comp if api_comp > 0 else orig_payg
        
        unaccounted = orig_payg - compute_payg
        win_lic_payg = 0
        if os_type == "Windows" and api_win > 0:
            win_lic_payg = min(max(0, unaccounted), api_win)
            unaccounted -= win_lic_payg
            
        prem_os_payg = 0
        sql_payg = 0
        if unaccounted > 5:
            if os_type in ["Red Hat", "SUSE"]:
                prem_os_payg = unaccounted
                unaccounted = 0
            elif sql_exact:
                sql_payg = unaccounted
                unaccounted = 0
            else:
                compute_payg += unaccounted 
                unaccounted = 0
                
        p["compute_payg_final"] = compute_payg
        p["win_lic_payg_final"] = win_lic_payg
        p["sql_payg_final"]     = sql_payg
        p["prem_os_payg_final"] = prem_os_payg
        
        if p.get("compute_ri1") is not None: p["compute_ri1"] *= qty
        if p.get("compute_ri3") is not None: p["compute_ri3"] *= qty

        row["api"] = p
        time.sleep(0.15) 

    return vm_rows

# ══════════════════════════════════════════════════════════════════════════════
#  OUTPUT 
# ══════════════════════════════════════════════════════════════════════════════
def write_res_header(ws):
    ws.merge_cells("F1:H1")
    hdr(ws["F1"], "Monthly Cost")
    for addr in ["G1","H1"]: ws[addr].border=BORDER
    for ci,h in enumerate(["Service category","Service type","Custom name", "Region","Description","PAYG", "1 Year RI Model","3 Year RI Model","Remarks"],1):
        hdr(ws.cell(2,ci), h, wrap=True)
    ws.row_dimensions[2].height = 28.8

def write_vm_sheet(wb, rows):
    ws = wb.create_sheet("Virtual Machines")
    write_res_header(ws)
    widths(ws,{"A":15,"B":15,"C":14,"D":12,"E":55,"F":13,"G":14,"H":14,"I":42})

    ri = 3
    total_payg = total_ri1 = total_ri3 = 0.0

    for row in rows:
        p = row.get("api", {})
        
        if p.get("is_standalone"):
            payg = row.get("payg", 0)
            vals = [row["svc_cat"], row["svc_type"], row["cust_name"], row["region"], row["desc"],
                    round(payg,2), round(payg,2), round(payg,2), row.get("remarks","")]
            for ci,v in enumerate(vals,1):
                dat(ws.cell(ri,ci), v, align="right" if ci>=6 and isinstance(v,float) else "left")
            ri += 1
            total_payg += payg
            total_ri1  += payg
            total_ri3  += payg
            continue

        compute_payg = p.get("compute_payg_final", row.get("payg",0))
        win_lic_payg = p.get("win_lic_payg_final", 0)
        prem_os_payg = p.get("prem_os_payg_final", 0)
        
        sql_rows_b = [s.copy() for s in row.get("sub_rows",[]) if "sql" in s["desc"].lower()]
        sql_payg_deduced = p.get("sql_payg_final", 0)
        
        if not sql_rows_b and sql_payg_deduced > 0:
            sql_rows_b.append({
                "desc": row.get("sql_lbl_exact", "SQL License"),
                "payg": sql_payg_deduced, "ri1": sql_payg_deduced, "ri3": sql_payg_deduced, "is_api": True
            })

        sql_payg_total = sum(s["payg"] for s in sql_rows_b if s.get("payg"))
        sql_ri1_total  = sum(s.get("ri1") or s["payg"] for s in sql_rows_b)
        sql_ri3_total  = sum(s.get("ri3") or s["payg"] for s in sql_rows_b)

        # Safe RI Fallback: Use exact API cost. If API failed entirely, gracefully default back to the baseline.
        api_ri1 = p.get("compute_ri1")
        compute_ri1 = api_ri1 if api_ri1 is not None else (row.get("ri1") if row.get("ri1") is not None else compute_payg)
        
        api_ri3 = p.get("compute_ri3")
        compute_ri3 = api_ri3 if api_ri3 is not None else (row.get("ri3") if row.get("ri3") is not None else compute_payg)
        
        # Strictly copy PAYG cost to RI columns for all licenses, NO zeroing out.
        prem_os_ri1 = prem_os_payg
        prem_os_ri3 = prem_os_payg

        vm_payg = compute_payg + win_lic_payg + prem_os_payg + sql_payg_total
        vm_ri1  = compute_ri1  + win_lic_payg + prem_os_ri1 + sql_ri1_total
        vm_ri3  = compute_ri3  + win_lic_payg + prem_os_ri3 + sql_ri3_total

        vals = [row["svc_cat"], row["svc_type"], row["cust_name"], row["region"], row["desc"],
                round(compute_payg,2), round(compute_ri1,2), round(compute_ri3,2), row.get("remarks","")]
        for ci,v in enumerate(vals,1):
            dat(ws.cell(ri,ci), v, align="right" if ci>=6 and isinstance(v,float) else "left")
        ri += 1

        if win_lic_payg > 0:
            sub = ["","","","","Windows License", round(win_lic_payg,2), round(win_lic_payg,2), round(win_lic_payg,2), "License Cost (Not discounted by Compute RI)"]
            for ci,v in enumerate(sub,1):
                dat(ws.cell(ri,ci), v, italic=True, color="595959", align="right" if ci>=6 and isinstance(v,float) else "left")
            ri += 1

        if prem_os_payg > 0:
            sub = ["","","","", row.get("os_lbl_exact", "Premium OS License"), round(prem_os_payg,2), round(prem_os_ri1,2), round(prem_os_ri3,2), "License Cost (Not discounted by Compute RI)"]
            for ci,v in enumerate(sub,1):
                dat(ws.cell(ri,ci), v, italic=True, color="595959", align="right" if ci>=6 and isinstance(v,float) else "left")
            ri += 1

        for s in sql_rows_b:
            rmk = "License Cost (Not discounted by Compute RI)" if s.get("is_api") else s.get("remarks", "")
            sub = ["","","","", s["desc"],
                   round(s["payg"],2) if s.get("payg") else None,
                   round(s.get("ri1") or s["payg"],2) if s.get("payg") else None,
                   round(s.get("ri3") or s["payg"],2) if s.get("payg") else None, rmk]
            for ci,v in enumerate(sub,1):
                dat(ws.cell(ri,ci), v, italic=True, color="595959", align="right" if ci>=6 and isinstance(v,float) else "left")
            ri += 1

        other_subs = [s for s in row.get("sub_rows",[]) if "sql" not in s["desc"].lower()]
        for s in other_subs:
            sub = ["","","","", s["desc"],
                   round(s["payg"],2) if s.get("payg") else None,
                   round(s.get("ri1") or s["payg"],2) if s.get("payg") else None,
                   round(s.get("ri3") or s["payg"],2) if s.get("payg") else None, s.get("remarks","")]
            for ci,v in enumerate(sub,1):
                dat(ws.cell(ri,ci), v, italic=True, color="595959", align="right" if ci>=6 and isinstance(v,float) else "left")
            ri += 1

        total_payg += vm_payg
        total_ri1  += vm_ri1
        total_ri3  += vm_ri3

    ws.cell(ri,5,"Total").font=_f(bold=True); ws.cell(ri,5).border=BORDER
    for ci,v in [(6,total_payg),(7,total_ri1),(8,total_ri3)]:
        tot(ws.cell(ri,ci), round(v,2))

    return total_payg, total_ri1, total_ri3

def write_generic_sheet(wb, sheet_name, rows):
    ws = wb.create_sheet(sheet_name)
    write_res_header(ws)
    widths(ws,{"A":15,"B":14,"C":22,"D":12,"E":60,"F":13,"G":14,"H":14,"I":40})

    ri=3; tp=tr1=tr3=0.0
    for row in rows:
        payg = row["payg"] or 0
        ri1, ri3 = row.get("ri1") or payg, row.get("ri3") or payg
        vals = [row["svc_cat"],row["svc_type"],row["cust_name"], row["region"],row["desc"], round(payg,2),round(ri1,2),round(ri3,2),row.get("remarks","")]
        for ci,v in enumerate(vals,1): dat(ws.cell(ri,ci),v, align="right" if ci>=6 and isinstance(v,float) else "left")
        tp+=payg; tr1+=ri1; tr3+=ri3; ri+=1

        for s in row.get("sub_rows",[]):
            sv=["","","","",s["desc"], round(s["payg"],2) if s.get("payg") else None, round(s.get("ri1") or s["payg"],2) if s.get("payg") else None, round(s.get("ri3") or s["payg"],2) if s.get("payg") else None, s.get("remarks","")]
            for ci,v in enumerate(sv,1): dat(ws.cell(ri,ci),v,italic=True,color="595959", align="right" if ci>=6 and isinstance(v,float) else "left")
            ri+=1

    ws.cell(ri,5,"Total").font=_f(bold=True); ws.cell(ri,5).border=BORDER
    for ci,v in [(6,tp),(7,tr1),(8,tr3)]: tot(ws.cell(ri,ci), round(v,2))
    return tp, tr1, tr3

def write_summary(wb, totals):
    ws = wb.create_sheet("Summary", 0)
    ws.merge_cells("A1:A2"); ws.merge_cells("B1:B2")
    ws.merge_cells("C1:E1"); ws.merge_cells("F1:F2")

    for addr,val,al in [("A1","Sl No","left"),("B1","Service Name","left"), ("C1","Monthly Cost","center"),("F1","Remarks","left")]:
        c=ws[addr]; c.value=val; c.font=_f(bold=True,size=11); c.alignment=_al(al,"center"); c.border=BORDER
    for addr in ["D1","E1","A2","B2","F2"]: ws[addr].border=BORDER
    for addr,val in [("C2","PAYG"),("D2","1 YR RI Model"),("E2","3 YR RI Model")]:
        c=ws[addr]; c.value=val; c.font=_f(bold=True,size=11); c.alignment=_al("center","center"); c.border=BORDER

    gp=gr1=gr3=0.0; ri=3
    for sl,(sname,(payg,ri1,ri3)) in enumerate(totals.items(),1):
        c=ws.cell(ri,1,sl); c.font=_f(size=11); c.border=BORDER; c.alignment=_al("center")
        c=ws.cell(ri,2,sname); c.font=_f(size=11); c.border=BORDER
        for ci,v in [(3,payg),(4,ri1),(5,ri3)]:
            c=ws.cell(ri,ci,round(v,2)); c.font=_f(size=11); c.alignment=_al("right"); c.border=BORDER; c.number_format=NUM
        ws.cell(ri,6).border=BORDER; gp+=payg; gr1+=ri1; gr3+=ri3; ri+=1

    for ci,v in [(3,gp),(4,gr1),(5,gr3)]: tot(ws.cell(ri,ci),round(v,2))
    for ci in [1,2,6]: ws.cell(ri,ci).border=BORDER
    widths(ws,{"A":5.33,"B":22,"C":13,"D":14,"E":14,"F":48})
    for r in range(1,ri+1): ws.row_dimensions[r].height=16.8

def convert(input_path, output_path, currency="INR"):
    print(f"\n{'='*62}")
    print(f"  Azure Cost Estimation Converter  (API-powered)")
    print(f"{'='*62}")
    print(f"  Input    : {input_path}")
    print(f"  Output   : {output_path}")
    print(f"  Currency : {currency}\n")

    wb_in = load_workbook(input_path, data_only=True)
    fmt = "A" if len(wb_in.sheetnames) == 1 else "B"
    rows = parse_format_a(wb_in) if fmt=="A" else parse_format_b(wb_in)
    
    if not rows:
        print("\n  ERROR: No data rows found. Check the file is an Azure Calculator export.")
        sys.exit(1)

    buckets = classify(rows)
    if "Virtual Machines" in buckets:
        enrich_vms(buckets["Virtual Machines"], currency)

    wb_out = Workbook(); wb_out.remove(wb_out.active); totals = {}

    for sname in SHEET_ORDER:
        if sname not in buckets: continue
        if sname == "Virtual Machines": p,r1,r3 = write_vm_sheet(wb_out, buckets[sname])
        else: p,r1,r3 = write_generic_sheet(wb_out, sname, buckets[sname])
        totals[sname]=(p,r1,r3)
        sav = f"  ({(p-r1)/p*100:.1f}% savings)" if p>0 and r1<p else ""
        print(f"  ✔  {sname:28s}  PAYG ₹{p:>12,.2f}   1YR ₹{r1:>12,.2f}   3YR ₹{r3:>12,.2f}{sav}")

    for sname,srows in buckets.items():
        if sname not in totals:
            p,r1,r3 = write_generic_sheet(wb_out, sname, srows)
            totals[sname]=(p,r1,r3)
            print(f"  ✔  {sname:28s}  PAYG ₹{p:>12,.2f}")

    write_summary(wb_out, totals)
    wb_out.save(output_path)
    
    gp, gr1, gr3  = sum(p for p,_,_ in totals.values()), sum(r for _,r,_ in totals.values()), sum(r for _,_,r in totals.values())
    sav1, sav3 = (f"{(gp-gr1)/gp*100:.1f}%" if gp>0 else "N/A"), (f"{(gp-gr3)/gp*100:.1f}%" if gp>0 else "N/A")
    print(f"\n  {'─'*56}")
    print(f"  Grand Total  PAYG ₹{gp:>12,.2f}   1YR ₹{gr1:>12,.2f}   3YR ₹{gr3:>12,.2f}")
    print(f"  Savings      1-Year: {sav1}            3-Year: {sav3}")
    print(f"  {'─'*56}\n  ✔ Output saved → {output_path}\n")

if __name__ == "__main__":
    args = sys.argv[1:]
    if not args:
        print('\n  Usage:   python convert.py "input file.xlsx"\n'); sys.exit(1)

    full = " ".join(args)
    
    # Safely find all ".xlsx" occurrences
    lower_full = full.lower()
    xlsx_positions = []
    pos = 0
    while True:
        idx = lower_full.find(".xlsx", pos)
        if idx == -1: break
        xlsx_positions.append(idx + 5)
        pos = idx + 1
        
    if not xlsx_positions:
        print(f'\n  ERROR: No .xlsx file found in: {full}\n')
        sys.exit(1)

    inp = full[:xlsx_positions[0]].strip()
    
    # Try to grab the currency code if provided at the end
    remainder = full[xlsx_positions[-1]:].strip()
    if remainder and len(remainder.split()[0]) == 3 and remainder.split()[0].isalpha():
        currency = remainder.split()[0].upper()
    else:
        currency = "INR"
        
    out = full[xlsx_positions[0]:xlsx_positions[1]].strip() if len(xlsx_positions) >= 2 else None

    inp_path = Path(inp)
    
    # Verbose Error Printing if file is missing
    if not inp_path.exists():
        print(f'\n  ERROR: File not found -> {inp}\n')
        print(f'  Please make sure you are running the script in the same folder')
        print(f'  as your Excel file, or provide the full path to the file.\n')
        print(f'  Current folder: {Path.cwd()}')
        print(f'  Available .xlsx files here:')
        for f in sorted(Path.cwd().glob("*.xlsx")):
            print(f'    {f.name}')
        print('\n')
        sys.exit(1)
        
    if out is None: out = str(inp_path.resolve().parent / "Cost_Estimation.xlsx")
    
    Path(out).parent.mkdir(parents=True, exist_ok=True)
    convert(inp, out, currency)
