"""
歴史的ショック対応モンテカルロシミュレーター
"""

import numpy as np
from typing import Dict, Tuple
from shock_analyzer import HistoricalShockAnalyzer

class ShockMonteCarloSimulator:
    """歴史的ショックを組み込んだモンテカルロシミュレーター"""
    
    def __init__(self, n_simulations: int = 10000):
        """
        Args:
            n_simulations: シミュレーション回数（固定10,000）
        """
        self.n_simulations = n_simulations
        self.shock_analyzer = HistoricalShockAnalyzer()
    
    def simulate_normal_scenario(self, params: Dict) -> Dict:
        """
        通常シナリオ（ショックなし）のシミュレーション
        
        Args:
            params: 基本パラメータ
            {
                'monthly_sales': 月次売上（万円）,
                'sales_volatility': 売上変動率（%）,
                'cash_balance': 現金残高（万円）,
                'cost_rate': 売上原価率（%）,
                'monthly_fixed_cost': 月次固定費（万円）,
                'ar_days': 売掛サイト（日）,
                'ap_days': 支払サイト（日）,
                'inventory_days': 在庫回転期間（日）
            }
            
        Returns:
            シミュレーション結果
            {
                'shortage_probability': 資金ショート確率（%）,
                'min_cash_mean': 最低残高平均（万円）,
                'min_cash_std': 最低残高標準偏差（万円）
            }
        """
        monthly_sales = params['monthly_sales']
        volatility = params['sales_volatility'] / 100
        cash = params['cash_balance']
        cost_rate = params['cost_rate'] / 100
        fixed_cost = params['monthly_fixed_cost']
        
        # 運転資金サイクル
        ar_days = params.get('ar_days', 30)
        ap_days = params.get('ap_days', 30)
        inventory_days = params.get('inventory_days', 0)
        
        # 12ヶ月シミュレーション
        n_months = 12
        shortage_count = 0
        min_cash_values = []
        
        for _ in range(self.n_simulations):
            cash_balance = cash
            min_cash = cash
            
            for month in range(n_months):
                # 売上生成（対数正規分布）
                sales = np.random.lognormal(
                    np.log(monthly_sales),
                    volatility
                )
                
                # キャッシュフロー計算
                cost = sales * cost_rate
                gross_profit = sales - cost
                operating_cf = gross_profit - fixed_cost
                
                # 運転資金の影響
                # 簡易版: (売掛-買掛-在庫) × 売上変動
                wc_impact = (ar_days - ap_days + inventory_days) / 30 * (sales - monthly_sales)
                
                cash_balance = cash_balance + operating_cf - wc_impact
                min_cash = min(min_cash, cash_balance)
            
            if min_cash < 0:
                shortage_count += 1
            
            min_cash_values.append(min_cash)
        
        shortage_prob = (shortage_count / self.n_simulations) * 100
        
        return {
            'shortage_probability': round(shortage_prob, 1),
            'min_cash_mean': round(np.mean(min_cash_values), 1),
            'min_cash_std': round(np.std(min_cash_values), 1)
        }
    
    def simulate_shock_scenario(self, params: Dict, industry: str, 
                                shock: str, service_sub: str = None) -> Dict:
        """
        歴史的ショックシナリオのシミュレーション
        
        Args:
            params: 基本パラメータ（simulate_normal_scenarioと同じ）
            industry: 業種名
            shock: ショック名（'リーマンショック'等）
            service_sub: サービス業サブカテゴリー
            
        Returns:
            シミュレーション結果（通常シナリオと同じ形式）
        """
        # ショックパラメータ取得
        shock_params = self.shock_analyzer.get_shock_params(
            industry, shock, service_sub
        )
        
        monthly_sales = params['monthly_sales']
        volatility = params['sales_volatility'] / 100
        cash = params['cash_balance']
        cost_rate = params['cost_rate'] / 100
        fixed_cost = params['monthly_fixed_cost']
        
        # ショックによる変化
        sales_decline = shock_params['sales_decline'] / 100
        cost_rate_increase = shock_params['cost_rate_increase'] / 100
        duration_months = shock_params['duration_months']
        
        # 運転資金サイクル
        ar_days = params.get('ar_days', 30)
        ap_days = params.get('ap_days', 30)
        inventory_days = params.get('inventory_days', 0)
        
        # 24ヶ月シミュレーション（ショック期間＋回復期間）
        n_months = 24
        shortage_count = 0
        min_cash_values = []
        
        for _ in range(self.n_simulations):
            cash_balance = cash
            min_cash = cash
            
            for month in range(n_months):
                # ショック影響の時間減衰
                if month < duration_months:
                    # ショック期間中
                    shock_factor = 1 + sales_decline
                    shock_cost_rate = cost_rate + cost_rate_increase
                else:
                    # 回復期間（線形回復）
                    recovery_progress = (month - duration_months) / 12  # 12ヶ月で回復
                    recovery_progress = min(recovery_progress, 1.0)
                    shock_factor = 1 + sales_decline * (1 - recovery_progress)
                    shock_cost_rate = cost_rate + cost_rate_increase * (1 - recovery_progress)
                
                # 売上生成（ショック影響込み）
                base_sales = np.random.lognormal(
                    np.log(monthly_sales),
                    volatility
                )
                sales = base_sales * shock_factor
                
                # キャッシュフロー計算
                cost = sales * shock_cost_rate
                gross_profit = sales - cost
                operating_cf = gross_profit - fixed_cost
                
                # 運転資金の影響
                wc_impact = (ar_days - ap_days + inventory_days) / 30 * (sales - monthly_sales * shock_factor)
                
                cash_balance = cash_balance + operating_cf - wc_impact
                min_cash = min(min_cash, cash_balance)
            
            if min_cash < 0:
                shortage_count += 1
            
            min_cash_values.append(min_cash)
        
        shortage_prob = (shortage_count / self.n_simulations) * 100
        
        return {
            'shortage_probability': round(shortage_prob, 1),
            'min_cash_mean': round(np.mean(min_cash_values), 1),
            'min_cash_std': round(np.std(min_cash_values), 1),
            'shock_severity': self.shock_analyzer.calculate_shock_severity(shock_params),
            'shock_description': shock_params['description']
        }
    
    def compare_all_shocks(self, params: Dict, industry: str, 
                          service_sub: str = None) -> Dict:
        """
        全ショックシナリオを比較
        
        Args:
            params: 基本パラメータ
            industry: 業種名
            service_sub: サービス業サブカテゴリー
            
        Returns:
            各ショックの結果辞書
            {
                '通常': {...},
                'リーマンショック': {...},
                ...
            }
        """
        results = {}
        
        # 通常シナリオ
        results['通常'] = self.simulate_normal_scenario(params)
        
        # 各ショックシナリオ
        for shock in HistoricalShockAnalyzer.AVAILABLE_SHOCKS:
            try:
                results[shock] = self.simulate_shock_scenario(
                    params, industry, shock, service_sub
                )
            except ValueError:
                # データがない場合はスキップ
                continue
        
        return results


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
    
    simulator = ShockMonteCarloSimulator(n_simulations=10000)
    
    # 製造業の全ショック比較
    print("=== 製造業の全ショック比較 ===")
    results = simulator.compare_all_shocks(test_params, '製造業')
    
    for scenario, result in results.items():
        print(f"\n【{scenario}】")
        print(f"資金ショート確率: {result['shortage_probability']}%")
        print(f"最低残高平均: {result['min_cash_mean']}万円")
        if 'shock_severity' in result:
            print(f"深刻度: {result['shock_severity']}")
            print(f"備考: {result['shock_description']}")
