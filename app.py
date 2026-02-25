
# ============================================================
# 資金繰りリスク可視化ツール - Streamlit WebUI (Phase 5)
# ============================================================
# 起動方法: streamlit run app.py
# ============================================================

import streamlit as st
import numpy as np
import matplotlib
matplotlib.use("Agg")
import matplotlib.pyplot as plt
import copy, io, warnings
from dataclasses import dataclass
from typing import Optional, List
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

warnings.filterwarnings("ignore")

st.set_page_config(
    page_title="資金繰りリスク可視化ツール",
    page_icon="📊", layout="wide",
    initial_sidebar_state="expanded")

# 日本語フォント設定（エラーが起きてもアプリを止めない）
import matplotlib.font_manager as fm, os as _os

try:
    _FONT_PATH = _os.path.join(_os.path.dirname(__file__), "ipag.ttf")
    if _os.path.exists(_FONT_PATH):
        fm.fontManager.addfont(_FONT_PATH)
        plt.rcParams["font.family"] = "IPAGothic"
    else:
        # システムフォントから日本語対応フォントを探す
        _JP_NAMES = {"IPAGothic","IPAPGothic","Noto Sans CJK JP","Noto Sans JP","MS Gothic","Yu Gothic","Meiryo"}
        _found = [f.name for f in fm.fontManager.ttflist if f.name in _JP_NAMES]
        if _found:
            plt.rcParams["font.family"] = _found[0]
except Exception:
    pass  # フォント設定失敗してもアプリは動かす
plt.rcParams["axes.unicode_minus"] = False

# ============================================================
# データクラス・マスタ
# ============================================================

@dataclass
class CashFlowParameters:
    monthly_sales: float; cash_balance: float
    monthly_fixed_cost: float; cost_rate: float
    sales_volatility: float = 0.15
    accounts_receivable_days: int = 30
    simulation_months: int = 12; num_simulations: int = 10000

    def gross_profit_rate(self):    return 1.0 - self.cost_rate
    def monthly_gross_profit(self): return self.monthly_sales * self.gross_profit_rate()
    def monthly_net_cf(self):       return self.monthly_gross_profit() - self.monthly_fixed_cost
    def breakeven_sales(self):      return self.monthly_fixed_cost / self.gross_profit_rate()
    def safety_months(self):        return self.cash_balance / self.monthly_fixed_cost

INDUSTRY_AR_MASTER = {
    "建設業":              {"standard":75,"floor":60,"note":"多重下請構造。個社努力での大幅短縮は困難。"},
    "製造業":              {"standard":52,"floor":45,"note":"系列・大企業主導。下請法の範囲内だが発注者が決定。"},
    "IT・システム開発":    {"standard":37,"floor":30,"note":"プロジェクト完了検収が起点。"},
    "広告・マーケティング":{"standard":75,"floor":60,"note":"代理店多段構造。直接取引比率向上が根本的改善。"},
    "卸売・小売":          {"standard":45,"floor":30,"note":"サプライチェーン上の立場で変動大。"},
    "EC・SaaS":           {"standard":15,"floor": 0,"note":"デジタル決済で既に短期化済。"},
    "医療・介護":          {"standard":45,"floor":30,"note":"保険診療分は短縮不可。自費部分は余地あり。"},
    "士業・コンサル":      {"standard":30,"floor":14,"note":"着手金・中間金の導入が最も現実的。"},
    "その他・不明":        {"standard":45,"floor":30,"note":"業種平均的な水準。"},
}

def get_industry_info(industry):
    return INDUSTRY_AR_MASTER.get(industry, INDUSTRY_AR_MASTER["その他・不明"])

# ============================================================
# シミュレーション関数
# ============================================================

def run_monte_carlo(params):
    n, T = params.num_simulations, params.simulation_months
    s = np.random.normal(params.monthly_sales,
                         params.monthly_sales * params.sales_volatility, (n, T))
    s = np.maximum(s, 0)
    cf = s * params.gross_profit_rate() - params.monthly_fixed_cost
    cash = np.zeros((n, T+1)); cash[:,0] = params.cash_balance
    for t in range(T): cash[:,t+1] = cash[:,t] + cf[:,t]
    return cash, s, cf

def calc_metrics(cash, sales_sim, params):
    bep = params.breakeven_sales()
    return {
        "shortage_prob":    float(np.mean(cash[:,1:].min(axis=1) < 0)),
        "monthly_shortage": np.mean(cash[:,1:] < 0, axis=0),
        "min_cash":         cash[:,1:].min(axis=1),
        "deficit_prob":     float(np.mean(sales_sim < bep)),
        "bep": bep,
        "p5":  np.percentile(cash[:,1:], 5, axis=0),
        "p25": np.percentile(cash[:,1:],25, axis=0),
        "p50": np.percentile(cash[:,1:],50, axis=0),
        "p75": np.percentile(cash[:,1:],75, axis=0),
        "p95": np.percentile(cash[:,1:],95, axis=0),
        "final_median": float(np.median(cash[:,-1])),
    }

def risk_icon(p):
    if p < 0.05: return "🟢 安全"
    if p < 0.15: return "🟡 注意"
    if p < 0.30: return "🟠 警戒"
    return "🔴 危険"

def risk_color(p):
    if p < 0.05: return "#27ae60"
    if p < 0.15: return "#f39c12"
    if p < 0.30: return "#e67e22"
    return "#e74c3c"

def run_stress(params, start_month, rate, duration=None):
    n, T = params.num_simulations, params.simulation_months
    s = np.random.normal(params.monthly_sales,
                         params.monthly_sales * params.sales_volatility, (n, T))
    s = np.maximum(s, 0)
    i0 = start_month - 1
    i1 = T if duration is None else min(i0 + duration, T)
    s[:,i0:i1] *= (1 - rate)
    cf = s * params.gross_profit_rate() - params.monthly_fixed_cost
    cash = np.zeros((n, T+1)); cash[:,0] = params.cash_balance
    for t in range(T): cash[:,t+1] = cash[:,t] + cf[:,t]
    return cash, s

def run_bankruptcy(params, month, receivable, share):
    n, T = params.num_simulations, params.simulation_months
    s = np.random.normal(params.monthly_sales,
                         params.monthly_sales * params.sales_volatility, (n, T))
    s = np.maximum(s, 0); s[:,month-1:] *= (1 - share)
    cf = s * params.gross_profit_rate() - params.monthly_fixed_cost
    cash = np.zeros((n, T+1)); cash[:,0] = params.cash_balance
    for t in range(T):
        cash[:,t+1] = cash[:,t] + cf[:,t]
        if t == month-1: cash[:,t+1] -= receivable
    return cash

# ============================================================
# 改善策エンジン
# ============================================================

@dataclass
class ImprovementMeasure:
    name: str; category: str; param_change: dict
    difficulty: str; difficulty_note: str; timeline: str; risk_note: str
    monthly_cost: float = 0.0

def build_measures(params, industry):
    measures = []
    ar=params.accounts_receivable_days; fc=params.monthly_fixed_cost; cr=params.cost_rate
    info=get_industry_info(industry); floor=info["floor"]; std=info["standard"]
    headroom=ar-floor
    if headroom <= 5:
        measures.append(ImprovementMeasure(
            name=f"売掛サイト短縮（業種下限{floor}日に近く実質困難）",
            category="receivable", param_change={}, difficulty="実質困難",
            difficulty_note=info["note"], timeline="—",
            risk_note="他の対策を優先してください。"))
    else:
        for cut, diff in [(max(headroom//3,3),"低" if ar>std else "低〜中"),
                          (headroom*2//3,"中"), (headroom,"高")]:
            if cut < 3: continue
            new_ar=ar-cut; release=(params.monthly_sales/30)*cut
            measures.append(ImprovementMeasure(
                name=f"売掛サイト {ar}日→{new_ar}日（{cut}日短縮）",
                category="receivable",
                param_change={"accounts_receivable_days":new_ar},
                difficulty=diff, difficulty_note=info["note"], timeline="3〜6ヶ月",
                risk_note=f"運転資本解放目安: 約{release:.0f}万円。"))
    for rate,diff,note in [
        (0.03,"低","通信費・消耗品・サブスク見直し。即着手できる。"),
        (0.05,"低〜中","リース契約見直し・外注費圧縮。"),
        (0.10,"中","賃料交渉・非正規雇用の調整が必要。")]:
        measures.append(ImprovementMeasure(
            name=f"固定費 {rate*100:.0f}%削減（月△{fc*rate:.0f}万円）",
            category="fixedcost",
            param_change={"monthly_fixed_cost":fc*(1-rate)},
            difficulty=diff, difficulty_note=note, timeline="1〜3ヶ月",
            risk_note="削減順序: 管理費→通信費→外注費→リース→人件費。"))
    for mp,diff,note in [
        (0.01,"低","一部商品・新規受注への価格改定。"),
        (0.02,"中","主力商品への価格改定が必要。"),
        (0.03,"高","全商品値上げまたは仕入先の大幅条件変更。")]:
        measures.append(ImprovementMeasure(
            name=f"粗利率 +{mp*100:.0f}pt（値上げ/仕入交渉）",
            category="margin",
            param_change={"cost_rate":max(cr-mp,0)},
            difficulty=diff, difficulty_note=note, timeline="3〜12ヶ月",
            risk_note="値上げしても販売数量が変わらない前提。"))
    R_OD,R_ST,R_LT = 0.025,0.0175,0.020
    for buf_m,diff,note in [(1,"低〜中","固定費1ヶ月分。"),(2,"中","固定費2ヶ月分。"),(3,"中","安全圏の3ヶ月分。")]:
        amount=fc*buf_m; interest=amount*R_OD/12
        measures.append(ImprovementMeasure(
            name=f"当座貸越枠 {amount:.0f}万円（固定費{buf_m}ヶ月分）",
            category="financing",
            param_change={"cash_balance":params.cash_balance+amount},
            difficulty=diff, difficulty_note=note, timeline="1〜2ヶ月",
            risk_note=f"引出時の月利息: {interest:.1f}万円（年率{R_OD*100:.1f}%）。",
            monthly_cost=interest))
    for loan_m in [3,6]:
        amount=fc*loan_m; repay=amount/12+amount*R_ST/12
        measures.append(ImprovementMeasure(
            name=f"短期借入 {amount:.0f}万円（固定費{loan_m}ヶ月・1年返済）",
            category="financing",
            param_change={"cash_balance":params.cash_balance+amount,
                          "monthly_fixed_cost":fc+repay},
            difficulty="中", difficulty_note="1年以内の返済を前提とした運転資金借入。",
            timeline="1〜2ヶ月",
            risk_note=f"月次返済（元利合計）: {repay:.1f}万円。",
            monthly_cost=repay))
    for years,diff,note in [(3,"中","3年返済。"),(5,"中","5年返済。"),(7,"中〜高","7年返済。")]:
        amount=fc*6; repay=amount/(years*12)+amount*R_LT/12
        measures.append(ImprovementMeasure(
            name=f"長期運転資金 {amount:.0f}万円（{years}年返済）",
            category="financing",
            param_change={"cash_balance":params.cash_balance+amount,
                          "monthly_fixed_cost":fc+repay},
            difficulty=diff, difficulty_note=note, timeline="2〜3ヶ月",
            risk_note=f"月次返済: {repay:.1f}万円/月。",
            monthly_cost=repay))
    return measures

def run_improvement_analysis(params, measures, base_prob):
    results = []
    for m in measures:
        p2 = copy.copy(params)
        for k, v in m.param_change.items(): setattr(p2, k, v)
        if m.category == "receivable" and m.param_change:
            cut = params.accounts_receivable_days - p2.accounts_receivable_days
            if cut > 0:
                p2.cash_balance = params.cash_balance + (params.monthly_sales/30)*cut
        np.random.seed(42)
        cash,s,_ = run_monte_carlo(p2)
        new_prob = float(np.mean(cash[:,1:].min(axis=1) < 0))
        results.append({"measure":m,"new_prob":new_prob,"improvement":base_prob-new_prob})
    results.sort(key=lambda x: -x["improvement"])
    return results

# ============================================================
# グラフ関数
# ============================================================

def fig_dashboard(cash, sales_sim, metrics, params):
    T = params.simulation_months
    months = np.arange(1, T+1)
    fig, axes = plt.subplots(2, 2, figsize=(14,10), facecolor="#f8f9fa")
    fig.suptitle(
        f"資金繰りリスクダッシュボード  ショート確率: {metrics['shortage_prob']:.1%}  {risk_icon(metrics['shortage_prob'])}",
        fontsize=13, fontweight="bold", y=0.98)

    ax1=axes[0,0]; ax1.set_facecolor("white")
    ax1.fill_between(months,metrics["p5"],metrics["p95"],alpha=0.12,color="steelblue",label="90%信頼区間")
    ax1.fill_between(months,metrics["p25"],metrics["p75"],alpha=0.28,color="steelblue",label="50%信頼区間")
    ax1.plot(months,metrics["p50"],"b-",lw=2.5,label="中央値")
    ax1.axhline(0,color="red",ls="--",lw=2,label="ショート境界線")
    ax1.set_title("現金残高の予測",fontweight="bold"); ax1.set_xlabel("月"); ax1.set_ylabel("万円")
    ax1.legend(fontsize=8); ax1.grid(alpha=0.3); ax1.set_xlim(0,T)

    ax2=axes[0,1]; ax2.set_facecolor("white")
    mp=metrics["monthly_shortage"]*100
    colors=["#2ecc71" if p<5 else "#f39c12" if p<15 else "#e67e22" if p<30 else "#e74c3c" for p in mp]
    ax2.bar(months,mp,color=colors,alpha=0.85,edgecolor="white")
    ax2.axhline(5,color="#f39c12",ls=":",lw=1.5,label="5%注意")
    ax2.axhline(15,color="#e67e22",ls=":",lw=1.5,label="15%警戒")
    ax2.set_title("月次ショート確率",fontweight="bold"); ax2.set_xlabel("月"); ax2.set_ylabel("%")
    ax2.legend(fontsize=8); ax2.grid(alpha=0.3,axis="y"); ax2.set_xlim(0.5,T+0.5)

    ax3=axes[1,0]; ax3.set_facecolor("white")
    mc=metrics["min_cash"]
    if (mc<0).any(): ax3.hist(mc[mc<0],bins=25,color="#e74c3c",alpha=0.7,label=f"ショート: {(mc<0).sum():,}回")
    ax3.hist(mc[mc>=0],bins=40,color="#3498db",alpha=0.7,label=f"安全: {(mc>=0).sum():,}回")
    ax3.axvline(0,color="red",ls="--",lw=2)
    ax3.set_title("最低現金残高の分布",fontweight="bold"); ax3.set_xlabel("万円"); ax3.set_ylabel("回数")
    ax3.legend(fontsize=8); ax3.grid(alpha=0.3)

    ax4=axes[1,1]; ax4.set_facecolor("white")
    bep=metrics["bep"]; all_s=sales_sim.flatten()
    ax4.hist(all_s[all_s>=bep],bins=50,color="#2ecc71",alpha=0.75,label=f"黒字 {(all_s>=bep).mean():.0%}")
    if (all_s<bep).any(): ax4.hist(all_s[all_s<bep],bins=30,color="#e74c3c",alpha=0.75,label=f"赤字 {(all_s<bep).mean():.0%}")
    ax4.axvline(bep,color="red",ls="--",lw=2,label=f"損益分岐点 {bep:,.0f}万円")
    ax4.set_title("損益分岐点分析",fontweight="bold"); ax4.set_xlabel("月次売上（万円）"); ax4.set_ylabel("頻度")
    ax4.legend(fontsize=8); ax4.grid(alpha=0.3)

    plt.tight_layout(rect=[0,0,1,0.96])
    return fig

def fig_comparison(results, title, params):
    names=[r[0] for r in results]; probs=[r[1]*100 for r in results]; finals=[r[2] for r in results]
    fig,(ax1,ax2)=plt.subplots(1,2,figsize=(14,6),facecolor="#f8f9fa")
    fig.suptitle(title,fontsize=13,fontweight="bold")

    bc=["#2ecc71" if p<5 else "#f1c40f" if p<15 else "#e67e22" if p<30 else "#e74c3c" for p in probs]
    bars=ax1.barh(range(len(names)),probs,color=bc,alpha=0.85,edgecolor="white",height=0.65)
    for bar,p in zip(bars,probs):
        ax1.text(bar.get_width()+0.3,bar.get_y()+bar.get_height()/2,f"{p:.1f}%",va="center",fontsize=8,fontweight="bold")
    for x,c,lbl in [(5,"#f39c12","5%注意"),(15,"#e67e22","15%警戒"),(30,"#e74c3c","30%危険")]:
        ax1.axvline(x,color=c,ls=":",lw=2,alpha=0.8,label=lbl)
    ax1.set_yticks(range(len(names))); ax1.set_yticklabels(names,fontsize=8)
    ax1.set_xlabel("ショート確率（%）"); ax1.set_title("ショートリスク比較",fontweight="bold")
    ax1.legend(fontsize=8); ax1.grid(alpha=0.3,axis="x"); ax1.set_facecolor("white")
    ax1.set_xlim(0,max(probs)*1.3+5)

    bc2=["#27ae60" if f>=500 else "#2ecc71" if f>=0 else "#e74c3c" for f in finals]
    bars2=ax2.barh(range(len(names)),finals,color=bc2,alpha=0.85,edgecolor="white",height=0.65)
    span=max(abs(min(finals)),max(finals)) if finals else 1
    for bar,f in zip(bars2,finals):
        ax2.text(bar.get_width()+span*0.02,bar.get_y()+bar.get_height()/2,f"{f:,.0f}",va="center",fontsize=8,fontweight="bold")
    ax2.axvline(0,color="red",ls="--",lw=2); ax2.axvline(params.cash_balance,color="navy",ls=":",lw=1.5)
    ax2.set_yticks(range(len(names))); ax2.set_yticklabels(names,fontsize=8)
    ax2.set_xlabel("12ヶ月後残高（中央値、万円）"); ax2.set_title("12ヶ月後の資金状況",fontweight="bold")
    ax2.grid(alpha=0.3,axis="x"); ax2.set_facecolor("white")

    plt.tight_layout()
    return fig

# ============================================================
# Excel出力
# ============================================================

FN="メイリオ"
CB="2C3E50"; CF="FFFFFF"; CS="D6EAF8"
CI="0000FF"; CG="F2F3F4"; CBR="BDC3C7"
CSAFE="27AE60"; CCAUT="F39C12"; CWARN="E67E22"; CDAN="E74C3C"

def _tb():
    s=Side(style="thin",color=CBR)
    return Border(left=s,right=s,top=s,bottom=s)

def _h(ws,row,col,val,w=None):
    c=ws.cell(row=row,column=col,value=val)
    c.font=Font(name=FN,bold=True,color=CF,size=10)
    c.fill=PatternFill("solid",start_color=CB)
    c.alignment=Alignment(horizontal="center",vertical="center",wrap_text=True)
    c.border=_tb()
    if w: ws.column_dimensions[get_column_letter(col)].width=w
    return c

def _s(ws,row,col,val):
    c=ws.cell(row=row,column=col,value=val)
    c.font=Font(name=FN,bold=True,size=10)
    c.fill=PatternFill("solid",start_color=CS)
    c.alignment=Alignment(horizontal="left",vertical="center")
    c.border=_tb(); return c

def _d(ws,row,col,val,fmt=None,bold=False,fg=None,bg=None,align="right"):
    c=ws.cell(row=row,column=col,value=val)
    c.font=Font(name=FN,bold=bold,color=fg or "000000",size=10)
    c.alignment=Alignment(horizontal=align,vertical="center")
    c.border=_tb()
    if fmt: c.number_format=fmt
    if bg: c.fill=PatternFill("solid",start_color=bg)
    return c

def _rc(p):
    if p<0.05: return CSAFE
    if p<0.15: return CCAUT
    if p<0.30: return CWARN
    return CDAN

def _rl(p):
    if p<0.05: return "安全"
    if p<0.15: return "注意"
    if p<0.30: return "警戒"
    return "危険"

def build_excel(params, industry, metrics,
                sens_results, stress_results, bankruptcy_results,
                interest_results, cost_results,
                improvement_results, base_prob):

    wb=Workbook(); wb.remove(wb.active)

    # ── Sheet1: 入力パラメータ ──
    ws=wb.create_sheet("①入力パラメータ"); ws.sheet_view.showGridLines=False
    ws.merge_cells("A1:D1"); t=ws["A1"]
    t.value="入力パラメータ一覧"
    t.font=Font(name=FN,bold=True,size=14,color=CF)
    t.fill=PatternFill("solid",start_color=CB)
    t.alignment=Alignment(horizontal="center",vertical="center")
    ws.row_dimensions[1].height=30
    ws.merge_cells("A2:D2"); ind=ws["A2"]
    ind.value=f"業種: {industry}"
    ind.font=Font(name=FN,bold=True,size=11)
    ind.fill=PatternFill("solid",start_color=CS)
    ind.alignment=Alignment(horizontal="left",vertical="center",indent=1)
    for col,(val,w) in enumerate([("項目",30),("入力値",15),("単位",10),("備考",30)],1):
        _h(ws,3,col,val,w)
    input_rows=[
        ("月次平均売上",params.monthly_sales,"万円","過去3〜6ヶ月の平均"),
        ("現金・預金残高",params.cash_balance,"万円","今日時点の通帳残高"),
        ("月次固定費",params.monthly_fixed_cost,"万円","家賃・人件費・リース等"),
        ("売上原価率",params.cost_rate,"","例: 60%なら0.60"),
        ("売上変動率",params.sales_volatility,"","不明なら0.15"),
        ("売掛サイト",params.accounts_receivable_days,"日","売上が現金になるまで"),
        ("シミュレーション期間",params.simulation_months,"ヶ月",""),
        ("シミュレーション回数",params.num_simulations,"回","多いほど精度UP"),
    ]
    calc_rows=[
        ("粗利率",params.gross_profit_rate(),"","1 − 原価率"),
        ("月次平均粗利",params.monthly_gross_profit(),"万円","売上 × 粗利率"),
        ("月次純CF（平均）",params.monthly_net_cf(),"万円","粗利 − 固定費"),
        ("損益分岐点",params.breakeven_sales(),"万円/月","固定費 ÷ 粗利率"),
        ("現金の安全余裕",params.safety_months(),"ヶ月分","現金 ÷ 固定費"),
    ]
    for i,(name,val,unit,note) in enumerate(input_rows,4):
        bg=CG if i%2==0 else "FFFFFF"
        _d(ws,i,1,name,align="left",bg=bg)
        fmt="0.0%" if unit=="" and val<=1 else "#,##0"
        _d(ws,i,2,val,fmt=fmt,bg=bg,fg=CI,bold=True)
        _d(ws,i,3,unit,align="center",bg=bg)
        _d(ws,i,4,note,align="left",bg=bg)
        ws.row_dimensions[i].height=20
    sep=len(input_rows)+4
    ws.merge_cells(f"A{sep}:D{sep}"); _s(ws,sep,1,"■ 自動計算値")
    for i,(name,val,unit,note) in enumerate(calc_rows,sep+1):
        bg=CG if i%2==0 else "FFFFFF"
        _d(ws,i,1,name,align="left",bg=bg)
        fmt="0.0%" if unit=="" and val<=1 else "#,##0.0" if "ヶ月" in unit else "#,##0"
        _d(ws,i,2,val,fmt=fmt,bg=bg)
        _d(ws,i,3,unit,align="center",bg=bg)
        _d(ws,i,4,note,align="left",bg=bg)
        ws.row_dimensions[i].height=20
    ws.freeze_panes="A4"

    # ── Sheet2: リスク診断 ──
    ws2=wb.create_sheet("②リスク診断結果"); ws2.sheet_view.showGridLines=False
    for col,w in [("A",28),("B",16),("C",12),("D",28)]: ws2.column_dimensions[col].width=w
    ws2.merge_cells("A1:D1"); t2=ws2["A1"]
    prob=metrics["shortage_prob"]
    t2.value="リスク診断結果"
    t2.font=Font(name=FN,bold=True,size=14,color=CF)
    t2.fill=PatternFill("solid",start_color=CB)
    t2.alignment=Alignment(horizontal="center",vertical="center")
    ws2.row_dimensions[1].height=30
    ws2.merge_cells("A2:D2"); r2=ws2["A2"]
    r2.value=f"総合判定: {_rl(prob)}  |  ショート確率: {prob:.1%}"
    r2.font=Font(name=FN,bold=True,size=13,color="FFFFFF")
    r2.fill=PatternFill("solid",start_color=_rc(prob))
    r2.alignment=Alignment(horizontal="center",vertical="center")
    ws2.row_dimensions[2].height=28
    for col,val in enumerate(["指標","値","単位","コメント"],1): _h(ws2,3,col,val)
    bep=metrics["bep"]; margin=(params.monthly_sales-bep)/params.monthly_sales
    rrows=[
        ("資金ショート確率",prob,"",f"判定: {_rl(prob)}"),
        ("12ヶ月後残高（中央値）",metrics["final_median"],"万円",""),
        ("損益分岐点（月次）",bep,"万円/月",""),
        ("安全余裕率",margin,"","プラスなら黒字基調"),
        ("赤字確率（平均）",metrics["deficit_prob"],"",""),
        ("現金の安全余裕",params.safety_months(),"ヶ月分","3ヶ月以上が目安"),
    ]
    for i,(name,val,unit,comment) in enumerate(rrows,4):
        bg=CG if i%2==0 else "FFFFFF"
        _d(ws2,i,1,name,align="left",bg=bg)
        fmt="0.0%" if unit=="" else "#,##0"
        vc=_d(ws2,i,2,val,fmt=fmt,bg=bg,bold=True)
        if name=="資金ショート確率":
            vc.fill=PatternFill("solid",start_color=_rc(prob))
            vc.font=Font(name=FN,bold=True,color="FFFFFF",size=10)
        _d(ws2,i,3,unit,align="center",bg=bg)
        _d(ws2,i,4,comment,align="left",bg=bg)
        ws2.row_dimensions[i].height=20
    sep2=len(rrows)+5
    ws2.merge_cells(f"A{sep2}:D{sep2}"); _s(ws2,sep2,1,"■ 月次ショート確率の推移")
    for col,val in enumerate(["月","ショート確率","判定",""],1): _h(ws2,sep2+1,col,val)
    for i,p in enumerate(metrics["monthly_shortage"],1):
        row=sep2+1+i; bg=CG if i%2==0 else "FFFFFF"
        _d(ws2,row,1,f"{i}ヶ月目",align="center",bg=bg)
        _d(ws2,row,2,p,fmt="0.0%",bg=bg,bold=True)
        _d(ws2,row,3,_rl(p),align="center",bg=_rc(p),fg="FFFFFF" if p>=0.05 else "000000")
        _d(ws2,row,4,"",bg=bg)
        ws2.row_dimensions[row].height=18
    ws2.freeze_panes="A4"

    # ── Sheet3: 感度分析 ──
    ws3=wb.create_sheet("③感度分析"); ws3.sheet_view.showGridLines=False
    for col,w in [("A",32),("B",16),("C",22),("D",14)]: ws3.column_dimensions[col].width=w
    ws3.merge_cells("A1:D1"); t3=ws3["A1"]
    t3.value="What-if 感度分析"
    t3.font=Font(name=FN,bold=True,size=14,color=CF)
    t3.fill=PatternFill("solid",start_color=CB)
    t3.alignment=Alignment(horizontal="center",vertical="center")
    ws3.row_dimensions[1].height=30
    for col,(val,w) in enumerate([("シナリオ",32),("ショート確率",16),("12ヶ月後残高（中央値）",22),("判定",12)],1):
        _h(ws3,2,col,val,w)
    ws3.row_dimensions[2].height=22
    for i,(name,p2,final) in enumerate(sens_results,3):
        bg=CG if i%2==1 else "FFFFFF"
        _d(ws3,i,1,name,align="left",bg=bg)
        _d(ws3,i,2,p2,fmt="0.0%",bg=bg,bold=True)
        _d(ws3,i,3,final,fmt="#,##0",bg=bg)
        _d(ws3,i,4,_rl(p2),align="center",bg=_rc(p2),
           fg="FFFFFF" if p2>=0.05 else "000000",bold=True)
        ws3.row_dimensions[i].height=20
    ws3.freeze_panes="A3"

    # ── Sheet4: ストレステスト ──
    ws4=wb.create_sheet("④ストレステスト"); ws4.sheet_view.showGridLines=False
    for col,w in [("A",38),("B",16),("C",20),("D",14)]: ws4.column_dimensions[col].width=w
    ws4.merge_cells("A1:D1"); t4=ws4["A1"]
    t4.value="ストレステスト結果"
    t4.font=Font(name=FN,bold=True,size=14,color=CF)
    t4.fill=PatternFill("solid",start_color=CB)
    t4.alignment=Alignment(horizontal="center",vertical="center")
    ws4.row_dimensions[1].height=30

    def write_block(ws,sr,title,results):
        ws.merge_cells(f"A{sr}:D{sr}"); _s(ws,sr,1,f"■ {title}"); ws.row_dimensions[sr].height=22
        for col,val in enumerate(["シナリオ","ショート確率","12ヶ月後残高（中央値）","判定"],1): _h(ws,sr+1,col,val)
        ws.row_dimensions[sr+1].height=22
        for i,(name,p2,final) in enumerate(results,sr+2):
            bg=CG if i%2==0 else "FFFFFF"
            _d(ws,i,1,name,align="left",bg=bg)
            _d(ws,i,2,p2,fmt="0.0%",bg=bg,bold=True)
            _d(ws,i,3,final,fmt="#,##0",bg=bg)
            _d(ws,i,4,_rl(p2),align="center",bg=_rc(p2),
               fg="FFFFFF" if p2>=0.05 else "000000",bold=True)
            ws.row_dimensions[i].height=20
        return sr+len(results)+3

    nr=write_block(ws4,2,"売上急落シナリオ",stress_results)
    nr=write_block(ws4,nr,"取引先倒産シナリオ",bankruptcy_results)
    nr=write_block(ws4,nr,"金利上昇シナリオ",interest_results)
    nr=write_block(ws4,nr,"仕入価格高騰シナリオ",cost_results)
    ws4.freeze_panes="A2"

    # ── Sheet5: 改善策提案 ──
    ws5=wb.create_sheet("⑤改善策提案"); ws5.sheet_view.showGridLines=False
    for col,w in [("A",38),("B",12),("C",12),("D",14),("E",14),("F",20),("G",16)]:
        ws5.column_dimensions[col].width=w
    ws5.merge_cells("A1:G1"); t5=ws5["A1"]
    t5.value=f"改善策提案  業種: {industry}  ベースライン: {base_prob:.1%}"
    t5.font=Font(name=FN,bold=True,size=13,color=CF)
    t5.fill=PatternFill("solid",start_color=CB)
    t5.alignment=Alignment(horizontal="center",vertical="center")
    ws5.row_dimensions[1].height=30
    DC={"低":"C8E6C9","低〜中":"FFF9C4","中":"FFE0B2",
        "中〜高":"FFCCBC","高":"FFCDD2","実質困難":"ECEFF1"}
    CJ={"receivable":"売掛サイト","fixedcost":"固定費","margin":"粗利率","financing":"資金調達"}
    for col,(val,w) in enumerate([
        ("改善策",38),("カテゴリ",12),("難易度",12),
        ("効果（pt）",14),("実施後確率",14),("実現期間",20),("月次コスト",16)],1):
        _h(ws5,2,col,val,w)
    ws5.row_dimensions[2].height=22
    for i,r in enumerate(improvement_results,3):
        m=r["measure"]; imp=r["improvement"]
        bg=DC.get(m.difficulty,"FFFFFF")
        _d(ws5,i,1,m.name,align="left",bg=bg)
        _d(ws5,i,2,CJ.get(m.category,m.category),align="center",bg=bg)
        _d(ws5,i,3,m.difficulty,align="center",bg=bg,bold=True)
        _d(ws5,i,4,-imp if imp>0.005 else 0,fmt="0.0%",bg=bg)
        _d(ws5,i,5,r["new_prob"],fmt="0.0%",bg=bg,bold=True)
        _d(ws5,i,6,m.timeline,align="left",bg=bg)
        cost=f"{m.monthly_cost:.1f}万円/月" if m.monthly_cost else "—"
        _d(ws5,i,7,cost,align="center",bg=bg)
        ws5.row_dimensions[i].height=20
    ws5.freeze_panes="A3"

    buf=io.BytesIO(); wb.save(buf); buf.seek(0)
    return buf

# ============================================================
# Streamlit UI
# ============================================================

def main():
    # ── サイドバー ──────────────────────────────────────────
    st.sidebar.title("📋 入力パラメータ")
    st.sidebar.markdown("---")

    industry=st.sidebar.selectbox("業種",list(INDUSTRY_AR_MASTER.keys()),index=2)
    info=get_industry_info(industry)
    st.sidebar.caption(f"売掛サイト 標準: {info['standard']}日 / 下限: {info['floor']}日")
    st.sidebar.markdown("---")

    monthly_sales=st.sidebar.number_input("① 月次平均売上（万円）",min_value=1,value=1000,step=50)
    cash_balance=st.sidebar.number_input("② 現金・預金残高（万円）",min_value=0,value=800,step=50)
    monthly_fixed_cost=st.sidebar.number_input("③ 月次固定費（万円）",min_value=1,value=320,step=10)
    cost_rate=st.sidebar.slider("④ 売上原価率（%）",min_value=0,max_value=99,value=60)/100
    sales_volatility=st.sidebar.slider("⑤ 売上変動率（%）",min_value=1,max_value=50,value=15)/100
    accounts_receivable_days=st.sidebar.number_input(
        "⑥ 売掛サイト（日）",min_value=0,max_value=180,value=info["standard"])
    simulation_months=st.sidebar.selectbox("⑦ シミュレーション期間",[ 6,12,18,24,36],index=1)
    num_simulations=st.sidebar.selectbox("⑧ シミュレーション回数",[1000,5000,10000],index=2)
    loan_balance=st.sidebar.number_input("借入金残高（万円）※金利上昇分析用",min_value=0,value=3000,step=100)
    current_rate=st.sidebar.number_input("現在の年利（%）※金利上昇分析用",min_value=0.0,max_value=10.0,value=1.0,step=0.1)/100

    st.sidebar.markdown("---")
    run_btn=st.sidebar.button("▶️ シミュレーション実行",type="primary",use_container_width=True)

    # ── メイン ──────────────────────────────────────────────
    st.title("📊 資金繰りリスク可視化ツール")
    st.caption("中小企業向け キャッシュフロー・リスク分析システム")

    if not run_btn:
        st.info("👈 左のサイドバーにパラメータを入力して「シミュレーション実行」を押してください。")
        st.markdown("""
        **このツールでできること**
        - 📊 Phase 1: 資金ショート確率シミュレーション・損益分岐点分析
        - 🔥 Phase 2: ストレステスト（売上急落・取引先倒産・金利上昇・仕入高騰）
        - 💡 Phase 3: 改善策提案エンジン（業種別ARモデル・資金調達オプション）
        - 📥 Phase 4: Excelダウンロード（5シート構成）
        """)

        # ── 利用マニュアルダウンロード ──
        st.markdown("---")
        st.markdown("#### 📖 利用マニュアル")
        import os as _os
        _manual_path = _os.path.join(_os.path.dirname(__file__), "資金繰りリスク可視化ツール_利用マニュアル.pdf")
        if _os.path.exists(_manual_path):
            with open(_manual_path, "rb") as _f:
                st.download_button(
                    label="📄 利用マニュアルをダウンロード（PDF）",
                    data=_f.read(),
                    file_name="資金繰りリスク可視化ツール_利用マニュアル.pdf",
                    mime="application/pdf",
                    use_container_width=True,
                )
        return

    # ── パラメータ構築 ──
    params=CashFlowParameters(
        monthly_sales=monthly_sales,cash_balance=cash_balance,
        monthly_fixed_cost=monthly_fixed_cost,cost_rate=cost_rate,
        sales_volatility=sales_volatility,
        accounts_receivable_days=accounts_receivable_days,
        simulation_months=simulation_months,num_simulations=num_simulations)

    # ============================================================
    # ★ 全シミュレーションをタブの外でまとめて実行 ★
    # ============================================================
    with st.spinner("⏳ シミュレーション実行中..."):
        np.random.seed(42)

        # Phase 1
        cash_paths,sales_sim,_=run_monte_carlo(params)
        metrics=calc_metrics(cash_paths,sales_sim,params)

        # 感度分析
        sens_scenarios={
            "①ベースケース（現状）":{},
            "②固定費10%削減":{"monthly_fixed_cost":params.monthly_fixed_cost*0.9},
            "③固定費20%削減":{"monthly_fixed_cost":params.monthly_fixed_cost*0.8},
            "④売上10%向上":{"monthly_sales":params.monthly_sales*1.1},
            "⑤売上10%減少":{"monthly_sales":params.monthly_sales*0.9},
            "⑥売上20%減少（ストレス）":{"monthly_sales":params.monthly_sales*0.8},
            "⑦変動リスク拡大（±30%）":{"sales_volatility":0.30},
            "⑧現金積増（+300万）":{"cash_balance":params.cash_balance+300},
        }
        sens_results=[]
        for name,changes in sens_scenarios.items():
            p2=copy.copy(params)
            for k,v in changes.items(): setattr(p2,k,v)
            np.random.seed(42)
            c2,s2,_=run_monte_carlo(p2)
            m2=calc_metrics(c2,s2,p2)
            sens_results.append((name,m2["shortage_prob"],m2["final_median"]))

        # ストレステスト（Phase 2）
        base_c,base_s,_=run_monte_carlo(params)

        stress_list=[("①ベースケース",base_c,base_s)]
        for label,rate,sm,dur in [
            ("②軽度（-20%・3ヶ月〜）",0.20,3,None),
            ("③中度（-30%・3ヶ月〜）",0.30,3,None),
            ("④重度（-50%・即時）",   0.50,1,None),
            ("⑤一時的ショック（-40%・3ヶ月間）",0.40,3,3)]:
            c2,s2=run_stress(params,start_month=sm,rate=rate,duration=dur)
            stress_list.append((label,c2,s2))
        stress_results=[(n,calc_metrics(c,s,params)["shortage_prob"],
                         float(np.median(c[:,-1]))) for n,c,s in stress_list]

        bkr_list=[("①ベースケース",base_c,base_s)]
        for label,recv,share in [
            ("②小規模（売上10%・売掛100万）",100,0.10),
            ("③中規模（売上20%・売掛300万）",300,0.20),
            ("④大口（売上30%・売掛500万）",  500,0.30),
            ("⑤最大手（売上40%・売掛800万）",800,0.40)]:
            c2=run_bankruptcy(params,month=6,receivable=recv,share=share)
            bkr_list.append((label,c2,base_s))
        bankruptcy_results=[(n,calc_metrics(c,s,params)["shortage_prob"],
                             float(np.median(c[:,-1]))) for n,c,s in bkr_list]

        int_list=[("①現状（金利変化なし）",base_c,base_s)]
        for new_rate,label in [
            (current_rate+0.005,"②金利+0.5%"),
            (current_rate+0.010,"③金利+1.0%"),
            (current_rate+0.020,"④金利+2.0%")]:
            extra=loan_balance*(new_rate-current_rate)/12
            p2=copy.copy(params); p2.monthly_fixed_cost+=extra
            c2,s2,_=run_monte_carlo(p2)
            int_list.append((label,c2,s2))
        interest_results=[(n,calc_metrics(c,s,params)["shortage_prob"],
                          float(np.median(c[:,-1]))) for n,c,s in int_list]

        cost_list=[("①現状（高騰なし）",base_c,base_s)]
        for inc,label in [(0.03,"②原価率+3pt"),(0.05,"③原価率+5pt"),(0.10,"④原価率+10pt")]:
            p2=copy.copy(params); p2.cost_rate=min(params.cost_rate+inc,0.99)
            c2,s2,_=run_monte_carlo(p2)
            cost_list.append((label,c2,s2))
        cost_results=[(n,calc_metrics(c,s,params)["shortage_prob"],
                      float(np.median(c[:,-1]))) for n,c,s in cost_list]

        # 改善策（Phase 3）
        cp_base,_,_=run_monte_carlo(params)
        base_prob=float(np.mean(cp_base[:,1:].min(axis=1)<0))
        measures=build_measures(params,industry)
        improvement_results=run_improvement_analysis(params,measures,base_prob)

    # ── KPIカード ──
    st.markdown("---")
    col1,col2,col3,col4=st.columns(4)
    col1.metric("資金ショート確率",f"{metrics['shortage_prob']:.1%}",delta=risk_icon(metrics['shortage_prob']))
    col2.metric("12ヶ月後残高（中央値）",f"{metrics['final_median']:,.0f}万円")
    col3.metric("損益分岐点（月次）",f"{metrics['bep']:,.0f}万円")
    col4.metric("現金の安全余裕",f"{params.safety_months():.1f}ヶ月分")

    # ── タブ ──
    tab1,tab2,tab3,tab4=st.tabs(
        ["📊 Phase 1: 基本分析","🔥 Phase 2: ストレステスト",
         "💡 Phase 3: 改善策提案","📥 Phase 4: Excelダウンロード"])

    with tab1:
        st.subheader("資金繰りリスクダッシュボード")
        fig1=fig_dashboard(cash_paths,sales_sim,metrics,params)
        st.pyplot(fig1); plt.close(fig1)
        st.subheader("What-if 感度分析")
        fig_s=fig_comparison(sens_results,"感度分析: シナリオ別リスク比較",params)
        st.pyplot(fig_s); plt.close(fig_s)

    with tab2:
        st.subheader("Phase 2-1: 売上急落ストレステスト")
        fig2=fig_comparison(stress_results,"売上急落ストレステスト",params)
        st.pyplot(fig2); plt.close(fig2)
        st.subheader("Phase 2-2: 取引先倒産シミュレーション")
        fig3=fig_comparison(bankruptcy_results,"取引先倒産シミュレーション",params)
        st.pyplot(fig3); plt.close(fig3)
        st.subheader("Phase 2-3: 金利上昇シミュレーション")
        fig4=fig_comparison(interest_results,f"金利上昇（借入{loan_balance:,}万円）",params)
        st.pyplot(fig4); plt.close(fig4)
        st.subheader("Phase 2-4: 仕入価格高騰シミュレーション")
        fig5=fig_comparison(cost_results,"仕入価格高騰シミュレーション",params)
        st.pyplot(fig5); plt.close(fig5)

    with tab3:
        st.subheader(f"改善策提案エンジン  業種: {industry}")
        st.metric("ベースライン ショート確率",f"{base_prob:.1%}")
        import pandas as pd
        CAT_JP={"receivable":"売掛サイト","fixedcost":"固定費",
                "margin":"粗利率","financing":"資金調達"}
        rows=[]
        for r in improvement_results:
            m=r["measure"]; imp=r["improvement"]
            rows.append({
                "改善策":m.name,
                "カテゴリ":CAT_JP.get(m.category,m.category),
                "難易度":m.difficulty,
                "効果（pt）":f"-{imp*100:.1f}pt" if imp>0.005 else "変化なし",
                "実施後確率":f"{r['new_prob']:.1%}",
                "実現期間":m.timeline,
                "月次コスト":f"{m.monthly_cost:.1f}万円/月" if m.monthly_cost else "—",
            })
        st.dataframe(pd.DataFrame(rows),use_container_width=True,hide_index=True)
        st.markdown("**💡 詳細コメント（上位5件）**")
        for r in improvement_results[:5]:
            m=r["measure"]
            with st.expander(f"{m.name}  難易度: {m.difficulty}"):
                st.write(f"**難易度根拠:** {m.difficulty_note}")
                st.write(f"**実現期間:** {m.timeline}")
                st.write(f"**留意事項:** {m.risk_note}")
                if m.monthly_cost: st.write(f"**月次コスト:** {m.monthly_cost:.1f}万円/月")

    with tab4:
        st.subheader("📥 Excelレポートのダウンロード")
        st.info("シミュレーション結果を5シート構成のExcelファイルに出力します。")
        excel_buf=build_excel(
            params,industry,metrics,
            sens_results,stress_results,bankruptcy_results,
            interest_results,cost_results,
            improvement_results,base_prob)
        st.download_button(
            label="📥 Excelをダウンロード",
            data=excel_buf,
            file_name="資金繰りリスク分析レポート.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            type="primary",use_container_width=True)
        st.markdown("""
        **【シート構成】**
        - ① 入力パラメータ
        - ② リスク診断結果
        - ③ 感度分析（What-if）
        - ④ ストレステスト
        - ⑤ 改善策提案
        """)

if __name__ == "__main__":
    main()
