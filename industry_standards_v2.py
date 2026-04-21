"""
業種×従業員数別 財務標準値マトリクス（サブカテゴリー対応版）
公開統計（中小企業実態基本調査・法人企業統計）ベース
"""

import pandas as pd
from typing import Dict, Optional

# 業種マッピング
INDUSTRY_MAP = {
    '建設業': 'construction',
    '製造業': 'manufacturing', 
    '卸・小売業': 'wholesale_retail',
    '運輸・物流業': 'transport_logistics',
    'サービス・その他': 'service_other'
}

# サービス業サブカテゴリー
SERVICE_SUBCATEGORIES = [
    'IT・システム開発',
    '医療・介護',
    '広告・マーケティング',
    'その他'
]

# 従業員数区分マッピング
EMPLOYEE_MAP = {
    '30人以下': '<=30',
    '30人超～50人以下': '31-50',
    '50人超～100人以下': '51-100',
    '100人超～300人以下': '101-300',
    '300人超': '>300'
}


class IndustryStandards:
    """業種別標準値を管理するクラス（サブカテゴリー対応版）"""
    
    def __init__(self, csv_path: str = 'industry_standards_matrix_v2.csv'):
        """
        Args:
            csv_path: 標準値マトリクスCSVのパス
        """
        self.df = pd.read_csv(csv_path, encoding='utf-8')
        
    def get_standards(self, industry: str, employee_count: str, 
                     service_sub: Optional[str] = None) -> Dict[str, float]:
        """
        指定した業種・従業員数の標準値を取得
        
        Args:
            industry: 業種名（例：'製造業'）
            employee_count: 従業員数区分（例：'30人以下'）
            service_sub: サービス業サブカテゴリー（サービス・その他の場合のみ必須）
            
        Returns:
            標準値の辞書
            {
                'cost_rate': 売上原価率（%）,
                'fixed_cost_rate': 固定費率（%）,
                'ar_days': 売掛サイト（日）,
                'ap_days': 支払サイト（日）,
                'inventory_days': 在庫回転期間（日）,
                'wip_days': 未成工事支出金回転期間（日）
            }
        """
        # サービス・その他の場合、サブカテゴリーが必須
        if industry == 'サービス・その他':
            if service_sub is None:
                raise ValueError("サービス・その他の場合、service_subパラメータが必要です")
            
            row = self.df[
                (self.df['業種'] == industry) & 
                (self.df['サブカテゴリー'] == service_sub) &
                (self.df['従業員数区分'] == employee_count)
            ]
        else:
            # サービス・その他以外はサブカテゴリー不要
            row = self.df[
                (self.df['業種'] == industry) & 
                (self.df['従業員数区分'] == employee_count)
            ]
        
        if row.empty:
            raise ValueError(f"データが見つかりません: {industry}, {employee_count}, {service_sub}")
        
        row = row.iloc[0]
        
        return {
            'cost_rate': row['売上原価率(%)'],
            'fixed_cost_rate': row['固定費率(%)'],
            'ar_days': row['売掛サイト(日)'],
            'ap_days': row['支払サイト(日)'],
            'inventory_days': row['在庫回転期間(日)'],
            'wip_days': row['未成工事支出金回転期間(日)'],
            'description': row['備考']
        }
    
    def get_benchmark_comparison(self, industry: str, employee_count: str, 
                                 actual_values: Dict[str, float],
                                 service_sub: Optional[str] = None) -> Dict[str, Dict]:
        """
        実績値と標準値の比較（偏差値計算）
        
        Args:
            industry: 業種名
            employee_count: 従業員数区分
            actual_values: 実績値の辞書（キーはget_standardsと同じ）
            service_sub: サービス業サブカテゴリー
            
        Returns:
            比較結果の辞書
        """
        standards = self.get_standards(industry, employee_count, service_sub)
        
        # 業種全体の標準偏差を取得
        if industry == 'サービス・その他' and service_sub:
            industry_data = self.df[
                (self.df['業種'] == industry) &
                (self.df['サブカテゴリー'] == service_sub)
            ]
        else:
            industry_data = self.df[self.df['業種'] == industry]
        
        result = {}
        
        for key in ['cost_rate', 'fixed_cost_rate', 'ar_days', 'ap_days']:
            if key not in actual_values:
                continue
                
            actual = actual_values[key]
            standard = standards[key]
            
            # 標準偏差を計算
            col_name = {
                'cost_rate': '売上原価率(%)',
                'fixed_cost_rate': '固定費率(%)',
                'ar_days': '売掛サイト(日)',
                'ap_days': '支払サイト(日)'
            }[key]
            
            std_dev = industry_data[col_name].std()
            
            # 標準偏差がゼロの場合（データが1つしかない）
            if std_dev == 0:
                std_dev = 1  # デフォルト値
            
            # 偏差値計算（原価率・固定費率・サイトは低いほど良い）
            if key in ['cost_rate', 'fixed_cost_rate', 'ar_days']:
                deviation = 50 - (actual - standard) / std_dev * 10
            else:  # ap_days（支払サイト）は長いほど良い
                deviation = 50 + (actual - standard) / std_dev * 10
            
            # 評価判定
            if deviation >= 55:
                evaluation = '良好'
            elif deviation >= 45:
                evaluation = '平均的'
            else:
                evaluation = '要改善'
            
            result[key] = {
                'actual': actual,
                'standard': standard,
                'deviation': round(deviation, 1),
                'evaluation': evaluation,
                'gap': round(actual - standard, 1)
            }
        
        return result
    
    def needs_inventory_input(self, industry: str) -> bool:
        """在庫回転期間の入力が必要な業種か判定"""
        return industry in ['製造業', '卸・小売業', '運輸・物流業']
    
    def needs_wip_input(self, industry: str) -> bool:
        """未成工事支出金回転期間の入力が必要な業種か判定"""
        return industry == '建設業'
    
    def needs_ar_input(self, industry: str, service_sub: Optional[str] = None) -> bool:
        """
        売掛サイトの入力が必要か判定
        
        Args:
            industry: 業種名
            service_sub: サービス業サブカテゴリー
            
        Returns:
            True: ユーザー入力必須
            False: 標準値使用可能
        """
        if industry != 'サービス・その他':
            # サービス業以外は常に入力必要
            return True
        
        # サービス業の場合、サブカテゴリーで判定
        if service_sub in ['IT・システム開発', '医療・介護', '広告・マーケティング']:
            return True
        else:
            return False
    
    def get_service_subcategories(self) -> list:
        """サービス業のサブカテゴリーリストを取得"""
        return SERVICE_SUBCATEGORIES


# 使用例
if __name__ == "__main__":
    standards = IndustryStandards('industry_standards_matrix_v2.csv')
    
    # サービス業以外の標準値取得
    print("=== 製造業・30人以下の標準値 ===")
    result = standards.get_standards('製造業', '30人以下')
    for key, value in result.items():
        print(f"{key}: {value}")
    
    # サービス業の標準値取得
    print("\n=== サービス・その他（IT・システム開発）・30人以下の標準値 ===")
    result_it = standards.get_standards('サービス・その他', '30人以下', 'IT・システム開発')
    for key, value in result_it.items():
        print(f"{key}: {value}")
    
    # 売掛サイト入力要否の判定
    print("\n=== 売掛サイト入力要否判定 ===")
    print(f"製造業: {standards.needs_ar_input('製造業')}")
    print(f"IT・システム開発: {standards.needs_ar_input('サービス・その他', 'IT・システム開発')}")
    print(f"その他サービス: {standards.needs_ar_input('サービス・その他', 'その他')}")
    
    # ベンチマーク比較（IT企業の例）
    print("\n=== ベンチマーク比較（IT企業・実績値が悪い例） ===")
    actual = {
        'cost_rate': 52,  # 標準48%より悪い
        'fixed_cost_rate': 36,  # 標準32%より悪い
        'ar_days': 85,  # 標準75日より悪い
        'ap_days': 22   # 標準25日より悪い
    }
    
    comparison = standards.get_benchmark_comparison(
        'サービス・その他', 
        '30人以下', 
        actual,
        'IT・システム開発'
    )
    for metric, data in comparison.items():
        print(f"\n{metric}:")
        print(f"  実績: {data['actual']}, 標準: {data['standard']}")
        print(f"  偏差値: {data['deviation']}, 評価: {data['evaluation']}")
        print(f"  乖離: {data['gap']}")
