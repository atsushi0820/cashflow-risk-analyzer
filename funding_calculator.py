"""
長期運転資金算出モジュール
目標資金ショート確率から必要資金額を逆算
"""

import numpy as np
from typing import Dict, List, Tuple

class FundingCalculator:
    """長期運転資金の必要額を算出するクラス"""
    
    # 目標水準の定義
    TARGET_LEVELS = {
        '安全': 5.0,   # 資金ショート確率5%以下
        '標準': 10.0,  # 資金ショート確率10%以下
        '最低限': 15.0  # 資金ショート確率15%以下
    }
    
    def __init__(self, n_simulations: int = 10000):
        """
        Args:
            n_simulations: シミュレーション回数
        """
        self.n_simulations = n_simulations
    
    def calculate_shortage_probability(self, params: Dict, 
                                      target_cash: float) -> float:
        """
        指定した現金残高での資金ショート確率を計算
        
        Args:
            params: 基本パラメータ
            target_cash: 目標現金残高（万円）
            
        Returns:
            資金ショート確率（%）
        """
        monthly_sales = params['monthly_sales']
        volatility = params['sales_volatility'] / 100
        cost_rate = params['cost_rate'] / 100
        fixed_cost = params['monthly_fixed_cost']
        
        ar_days = params.get('ar_days', 30)
        ap_days = params.get('ap_days', 30)
        inventory_days = params.get('inventory_days', 0)
        
        n_months = 12
        shortage_count = 0
        
        for _ in range(self.n_simulations):
            cash_balance = target_cash
            
            for month in range(n_months):
                # 売上生成
                sales = np.random.lognormal(
                    np.log(monthly_sales),
                    volatility
                )
                
                # キャッシュフロー計算
                cost = sales * cost_rate
                gross_profit = sales - cost
                operating_cf = gross_profit - fixed_cost
                
                # 運転資金の影響
                wc_impact = (ar_days - ap_days + inventory_days) / 30 * (sales - monthly_sales)
                
                cash_balance = cash_balance + operating_cf - wc_impact
            
            # 12ヶ月間で一度でもマイナスになったかチェック
            if cash_balance < 0:
                shortage_count += 1
        
        return (shortage_count / self.n_simulations) * 100
    
    def find_required_cash(self, params: Dict, target_probability: float,
                          tolerance: float = 0.5) -> float:
        """
        二分探索で目標確率を達成する現金残高を算出
        
        Args:
            params: 基本パラメータ
            target_probability: 目標資金ショート確率（%）
            tolerance: 許容誤差（%）
            
        Returns:
            必要現金残高（万円）
        """
        # 初期範囲設定
        low = 0
        high = params['monthly_sales'] * 24  # 最大月商の24倍まで
        
        # 二分探索
        max_iterations = 20
        for i in range(max_iterations):
            mid = (low + high) / 2
            
            prob = self.calculate_shortage_probability(params, mid)
            
            # 目標確率に到達したか確認
            if abs(prob - target_probability) < tolerance:
                return mid
            
            if prob > target_probability:
                # 確率が高すぎる → 現金残高を増やす
                low = mid
            else:
                # 確率が低すぎる → 現金残高を減らす
                high = mid
        
        # 収束しない場合は最終値を返す
        return mid
    
    def calculate_funding_needs(self, params: Dict) -> Dict:
        """
        全目標水準の必要資金額を算出
        
        Args:
            params: 基本パラメータ
            
        Returns:
            目標水準別の必要資金額辞書
            {
                '安全': {
                    'target_probability': 5.0,
                    'required_cash': 1600.0,
                    'funding_amount': 800.0,
                    ...
                },
                ...
            }
        """
        current_cash = params['cash_balance']
        results = {}
        
        for level, target_prob in self.TARGET_LEVELS.items():
            # 必要現金残高を算出
            required_cash = self.find_required_cash(params, target_prob)
            
            # 必要資金額（現状との差額）
            funding_amount = max(0, required_cash - current_cash)
            
            # 月次返済額（5年固定）
            monthly_repayment_5y = funding_amount / 60 if funding_amount > 0 else 0
            
            results[level] = {
                'target_probability': target_prob,
                'required_cash': round(required_cash, 1),
                'funding_amount': round(funding_amount, 1),
                'monthly_repayment_5y': round(monthly_repayment_5y, 1)
            }
        
        return results
    
    def calculate_monthly_cf_surplus(self, params: Dict) -> float:
        """
        月次CF余剰を計算
        
        Args:
            params: 基本パラメータ
            
        Returns:
            月次CF余剰（万円）
        """
        monthly_sales = params['monthly_sales']
        cost_rate = params['cost_rate'] / 100
        fixed_cost = params['monthly_fixed_cost']
        
        # 平均的な営業CF
        gross_profit = monthly_sales * (1 - cost_rate)
        operating_cf = gross_profit - fixed_cost
        
        # 既存借入返済額（パラメータにあれば）
        existing_repayment = params.get('existing_debt_repayment', 0)
        
        cf_surplus = operating_cf - existing_repayment
        
        return cf_surplus
    
    def calculate_shortest_repayment(self, funding_amount: float, 
                                    monthly_cf_surplus: float,
                                    safety_margin: float = 0.2) -> Tuple[float, float]:
        """
        最短償還年数を計算
        
        Args:
            funding_amount: 必要資金額（万円）
            monthly_cf_surplus: 月次CF余剰（万円）
            safety_margin: 安全マージン（0.2 = 20%を安全マージンとして確保）
            
        Returns:
            (最短償還年数, 月次返済額) のタプル
        """
        if monthly_cf_surplus <= 0 or funding_amount <= 0:
            return 0, 0
        
        # CF余剰の一定割合を返済に充当
        max_monthly_repayment = monthly_cf_surplus * (1 - safety_margin)
        
        # 最短償還月数
        shortest_months = funding_amount / max_monthly_repayment
        shortest_years = shortest_months / 12
        
        return round(shortest_years, 1), round(max_monthly_repayment, 1)
    
    def generate_repayment_plans(self, funding_amount: float,
                                monthly_cf_surplus: float) -> List[Dict]:
        """
        複数の返済プランを生成
        
        Args:
            funding_amount: 必要資金額（万円）
            monthly_cf_surplus: 月次CF余剰（万円）
            
        Returns:
            返済プランのリスト
        """
        if funding_amount <= 0:
            return []
        
        plans = []
        
        # 標準返済（5年）
        monthly_5y = funding_amount / 60
        margin_5y = monthly_cf_surplus - monthly_5y
        
        plans.append({
            'name': '標準返済（推奨）',
            'years': 5,
            'months': 60,
            'monthly_repayment': round(monthly_5y, 1),
            'cf_margin': round(margin_5y, 1),
            'evaluation': self._evaluate_plan(margin_5y)
        })
        
        # 短期返済（3年）
        monthly_3y = funding_amount / 36
        margin_3y = monthly_cf_surplus - monthly_3y
        
        plans.append({
            'name': '短期返済',
            'years': 3,
            'months': 36,
            'monthly_repayment': round(monthly_3y, 1),
            'cf_margin': round(margin_3y, 1),
            'evaluation': self._evaluate_plan(margin_3y)
        })
        
        # 最短返済（参考）
        shortest_years, max_repayment = self.calculate_shortest_repayment(
            funding_amount, monthly_cf_surplus
        )
        
        if shortest_years > 0:
            margin_shortest = monthly_cf_surplus - max_repayment
            
            plans.append({
                'name': '最短返済（参考）',
                'years': shortest_years,
                'months': round(shortest_years * 12),
                'monthly_repayment': max_repayment,
                'cf_margin': round(margin_shortest, 1),
                'evaluation': self._evaluate_plan(margin_shortest)
            })
        
        return plans
    
    def _evaluate_plan(self, cf_margin: float) -> str:
        """
        返済プランの評価
        
        Args:
            cf_margin: CF余裕（万円）
            
        Returns:
            評価文字列
        """
        if cf_margin >= 10:
            return '✅ 余裕あり'
        elif cf_margin >= 5:
            return '⚠️ やや厳しい'
        else:
            return '❌ 余裕なし'


# 使用例
if __name__ == "__main__":
    # テスト用パラメータ
    test_params = {
        'monthly_sales': 1000,
        'sales_volatility': 15,
        'cash_balance': 800,
        'cost_rate': 65,
        'monthly_fixed_cost': 320,
        'ar_days': 45,
        'ap_days': 35,
        'inventory_days': 30
    }
    
    calculator = FundingCalculator(n_simulations=10000)
    
    # 必要資金額の算出
    print("=== 長期運転資金の必要額算出 ===\n")
    funding_needs = calculator.calculate_funding_needs(test_params)
    
    for level, result in funding_needs.items():
        print(f"【{level}水準】")
        print(f"  目標確率: {result['target_probability']}%以下")
        print(f"  必要現金残高: {result['required_cash']:,.1f}万円")
        print(f"  必要資金額: {result['funding_amount']:,.1f}万円")
        print(f"  月次返済額（5年）: {result['monthly_repayment_5y']:,.1f}万円")
        print()
    
    # 返済プランの生成
    print("\n=== 返済プラン比較（安全水準の場合） ===\n")
    
    funding_amount = funding_needs['安全']['funding_amount']
    monthly_cf_surplus = calculator.calculate_monthly_cf_surplus(test_params)
    
    print(f"必要資金額: {funding_amount:,.1f}万円")
    print(f"月次CF余剰: {monthly_cf_surplus:,.1f}万円\n")
    
    plans = calculator.generate_repayment_plans(funding_amount, monthly_cf_surplus)
    
    for plan in plans:
        print(f"【{plan['name']}】")
        print(f"  返済期間: {plan['years']}年（{plan['months']}ヶ月）")
        print(f"  月次返済額: {plan['monthly_repayment']:,.1f}万円")
        print(f"  CF余裕: {plan['cf_margin']:,.1f}万円")
        print(f"  評価: {plan['evaluation']}")
        print()
