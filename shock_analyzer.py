"""
歴史的ストレステスト分析モジュール
リーマンショック、東日本大震災、コロナショック、ウクライナ戦争
"""

import pandas as pd
import numpy as np
from typing import Dict, List, Tuple, Optional

class HistoricalShockAnalyzer:
    """歴史的ショックシナリオを管理・分析するクラス"""
    
    # 利用可能なショックシナリオ
    AVAILABLE_SHOCKS = [
        'リーマンショック',
        '東日本大震災',
        'コロナショック',
        'ウクライナ戦争'
    ]
    
    def __init__(self, csv_path: str = 'shock_scenarios.csv'):
        """
        Args:
            csv_path: ショックシナリオCSVのパス
        """
        self.df = pd.read_csv(csv_path, encoding='utf-8')
    
    def get_shock_params(self, industry: str, shock: str, 
                        service_sub: Optional[str] = None) -> Dict:
        """
        指定した業種・ショックのパラメータを取得
        
        Args:
            industry: 業種名（例：'製造業'）
            shock: ショック名（例：'リーマンショック'）
            service_sub: サービス業サブカテゴリー
            
        Returns:
            ショックパラメータの辞書
            {
                'sales_decline': 売上減少率（%）,
                'cost_rate_increase': 原価率上昇（%pt）,
                'duration_months': 影響期間（月）,
                'bankruptcy_prob_increase': 倒産確率上昇（%pt）,
                'description': 備考
            }
        """
        # サービス・その他の場合、サブカテゴリーで絞り込み
        if industry == 'サービス・その他':
            if service_sub is None:
                raise ValueError("サービス・その他の場合、service_subパラメータが必要です")
            
            row = self.df[
                (self.df['業種'] == industry) & 
                (self.df['サブカテゴリー'] == service_sub) &
                (self.df['ショック'] == shock)
            ]
        else:
            row = self.df[
                (self.df['業種'] == industry) & 
                (self.df['ショック'] == shock)
            ]
        
        if row.empty:
            raise ValueError(f"データが見つかりません: {industry}, {shock}, {service_sub}")
        
        row = row.iloc[0]
        
        return {
            'sales_decline': row['売上減少率(%)'],
            'cost_rate_increase': row['原価率上昇(%pt)'],
            'duration_months': int(row['影響期間(月)']),
            'bankruptcy_prob_increase': row['倒産確率上昇(%pt)'],
            'description': row['備考']
        }
    
    def compare_all_shocks(self, industry: str, 
                          service_sub: Optional[str] = None) -> pd.DataFrame:
        """
        全ショックシナリオを比較
        
        Args:
            industry: 業種名
            service_sub: サービス業サブカテゴリー
            
        Returns:
            比較表のDataFrame
        """
        results = []
        
        for shock in self.AVAILABLE_SHOCKS:
            try:
                params = self.get_shock_params(industry, shock, service_sub)
                results.append({
                    'ショック': shock,
                    '売上減少率(%)': params['sales_decline'],
                    '原価率上昇(%pt)': params['cost_rate_increase'],
                    '影響期間(月)': params['duration_months'],
                    '倒産確率上昇(%pt)': params['bankruptcy_prob_increase'],
                    '備考': params['description']
                })
            except ValueError:
                # データがない場合はスキップ
                continue
        
        return pd.DataFrame(results)
    
    def get_worst_shock(self, industry: str, 
                       service_sub: Optional[str] = None) -> Tuple[str, Dict]:
        """
        最悪ケースのショックを特定
        
        Args:
            industry: 業種名
            service_sub: サービス業サブカテゴリー
            
        Returns:
            (ショック名, パラメータ辞書) のタプル
        """
        comparison = self.compare_all_shocks(industry, service_sub)
        
        # 総合影響度 = 売上減少率 × 影響期間 / 12
        comparison['総合影響度'] = (
            abs(comparison['売上減少率(%)']) * 
            comparison['影響期間(月)'] / 12
        )
        
        worst_idx = comparison['総合影響度'].idxmax()
        worst_shock = comparison.loc[worst_idx, 'ショック']
        
        params = self.get_shock_params(industry, worst_shock, service_sub)
        
        return worst_shock, params
    
    def calculate_shock_severity(self, shock_params: Dict) -> str:
        """
        ショックの深刻度を判定
        
        Args:
            shock_params: get_shock_paramsの戻り値
            
        Returns:
            '甚大'/'大'/'中'/'小'
        """
        sales_decline = abs(shock_params['sales_decline'])
        duration = shock_params['duration_months']
        
        # 総合影響度 = 売上減少率 × 影響期間 / 12
        severity_score = sales_decline * duration / 12
        
        if severity_score >= 30:
            return '甚大'
        elif severity_score >= 20:
            return '大'
        elif severity_score >= 10:
            return '中'
        else:
            return '小'
    
    def get_industry_vulnerability_matrix(self) -> pd.DataFrame:
        """
        業種別×ショック別の脆弱性マトリクスを作成
        
        Returns:
            脆弱性マトリクスのDataFrame
        """
        # 業種とショックの組み合わせごとに総合影響度を計算
        matrix_data = []
        
        industries = self.df['業種'].unique()
        
        for industry in industries:
            if industry == 'サービス・その他':
                # サブカテゴリー別に計算
                subcategories = self.df[
                    self.df['業種'] == 'サービス・その他'
                ]['サブカテゴリー'].unique()
                
                for sub in subcategories:
                    row_data = {'業種': f"{industry}（{sub}）"}
                    
                    for shock in self.AVAILABLE_SHOCKS:
                        try:
                            params = self.get_shock_params(industry, shock, sub)
                            severity_score = (
                                abs(params['sales_decline']) * 
                                params['duration_months'] / 12
                            )
                            row_data[shock] = round(severity_score, 1)
                        except ValueError:
                            row_data[shock] = 0
                    
                    matrix_data.append(row_data)
            else:
                row_data = {'業種': industry}
                
                for shock in self.AVAILABLE_SHOCKS:
                    try:
                        params = self.get_shock_params(industry, shock)
                        severity_score = (
                            abs(params['sales_decline']) * 
                            params['duration_months'] / 12
                        )
                        row_data[shock] = round(severity_score, 1)
                    except ValueError:
                        row_data[shock] = 0
                
                matrix_data.append(row_data)
        
        return pd.DataFrame(matrix_data)


# 使用例
if __name__ == "__main__":
    analyzer = HistoricalShockAnalyzer('shock_scenarios.csv')
    
    # 製造業のショック比較
    print("=== 製造業の全ショック比較 ===")
    comparison = analyzer.compare_all_shocks('製造業')
    print(comparison.to_string(index=False))
    
    # 最悪ケースの特定
    print("\n=== 製造業の最悪ケース ===")
    worst_shock, params = analyzer.get_worst_shock('製造業')
    severity = analyzer.calculate_shock_severity(params)
    print(f"最悪ショック: {worst_shock}")
    print(f"売上減少率: {params['sales_decline']}%")
    print(f"影響期間: {params['duration_months']}ヶ月")
    print(f"深刻度: {severity}")
    
    # サービス業（IT）のショック比較
    print("\n=== サービス業（IT・システム開発）の全ショック比較 ===")
    comparison_it = analyzer.compare_all_shocks('サービス・その他', 'IT・システム開発')
    print(comparison_it.to_string(index=False))
    
    # 業種別脆弱性マトリクス
    print("\n=== 業種別×ショック別 脆弱性マトリクス ===")
    matrix = analyzer.get_industry_vulnerability_matrix()
    print(matrix.to_string(index=False))
    print("\n※数値は総合影響度（売上減少率×影響期間/12）")
