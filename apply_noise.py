import numpy as np
import random
from typing import Dict, List, Optional

class GradeReverseEngine:
    """
    课程成绩逆向工程生成器 (GradeReverseEngine)
    功能：
    1. 根据总分逆向生成明细分。
    2. 支持多种统计分布 (正态、左偏、双峰等)。
    3. 支持【噪声三角】控制：触发率、分数分布模式、位置偏好。
    """

    def __init__(self):
        pass

    # ==========================================
    # 1. 核心分布生成方法 (Distribution Methods)
    # ==========================================

    def dist_normal(self, target_mean: float, scale: float = 5.0) -> float:
        """标准正态分布"""
        score = np.random.normal(loc=target_mean, scale=scale)
        return self._clamp(score)

    def dist_left_skewed(self, target_mean: float, strength: float = 10.0) -> float:
        """左偏分布 (容易拿高分)"""
        low = max(0, target_mean - strength * 3)
        high = min(100, target_mean + strength)
        mode = high 
        score = np.random.triangular(left=low, mode=mode, right=high)
        return self._clamp(score)

    def dist_right_skewed(self, target_mean: float, strength: float = 10.0) -> float:
        """右偏分布 (题目难，低分多)"""
        low = max(0, target_mean - strength)
        high = min(100, target_mean + strength * 3)
        mode = low 
        score = np.random.triangular(left=low, mode=mode, right=high)
        return self._clamp(score)

    def dist_bimodal(self, low_peak: float = 60, high_peak: float = 90, ratio: float = 0.5) -> float:
        """双峰分布 (两极分化)"""
        if random.random() < ratio:
            score = np.random.normal(loc=high_peak, scale=5)
        else:
            score = np.random.normal(loc=low_peak, scale=5)
        return self._clamp(score)

    def dist_discrete(self, levels: List[int] = [60, 70, 80, 85, 90, 95]) -> float:
        """离散档位分布"""
        score = np.random.choice(levels)
        return float(score)

    # ==========================================
    # 2. 智能噪声控制 (Smart Noise Injection)
    # ==========================================

    def apply_advanced_noise(self, 
                             scores_map: Dict[str, float], 
                             noise_ratio: float,          # 控制1：触发率 (0-1)
                             severity_mode: str,          # 控制2：数值分布模式
                             allowed_items: List[str]     # 控制3：位置偏好 (白名单)
                             ) -> Dict[str, float]:
        """
        注入高级噪声
        :param noise_ratio: 学生出现不及格项的概率 (0.0 - 1.0)
        :param severity_mode: 'random'(40-59), 'near_miss'(55-59), 'catastrophic'(0-40)
        :param allowed_items: 允许注入噪声的科目名称列表
        """
        noisy_scores = scores_map.copy()
        
        # --- 控制1：决定这个学生是否中招 ---
        # 如果随机数大于比率，则该学生安全，直接返回原成绩
        if random.random() > noise_ratio:
            return noisy_scores 

        # --- 控制3：决定在哪一项注入 ---
        # 过滤出当前学生有的、且在允许列表里的科目
        valid_targets = [k for k in noisy_scores.keys() if k in allowed_items]
        
        # 如果没有合法的注入目标（比如只剩考勤了，但考勤不允许挂科），则跳过
        if not valid_targets:
            return noisy_scores

        # 随机挑一个倒霉的科目
        target_item = random.choice(valid_targets)
        
        # --- 控制2：决定不及格的分数是多少 (核心修改区域) ---
        if severity_mode == 'near_miss': 
            # 边缘挂科: 55 - 59 分 (模拟努力了但没过)
            fake_score = random.uniform(55, 59.9)
            
        elif severity_mode == 'catastrophic': 
            # 严重缺失: 0 - 40 分 (模拟缺考或极差)
            fake_score = random.uniform(0, 40.0)
            
        else: 
            # random / default: 40 - 59 分 (常规不及格)
            # 您的要求：默认推荐改为 40-59
            fake_score = random.uniform(40, 59.9)
            
        # 注入分数 (保留1位小数)
        noisy_scores[target_item] = round(fake_score, 1)
        
        return noisy_scores

    # ==========================================
    # 3. 核心逆向引擎 (Main Generator)
    # ==========================================

    def generate_breakdown(self, 
                           student_total_score: float, 
                           structure: Dict[str, Dict], 
                           noise_config: Dict = None) -> Dict[str, float]:
        """
        主函数：根据总分逆向生成各分项
        :param noise_config: 包含 noise_ratio, severity_mode, allowed_items 的字典
        """
        
        # 默认噪声配置
        if noise_config is None:
            noise_config = {
                'noise_ratio': 0.0, 
                'severity_mode': 'random', 
                'allowed_items': []
            }

        # 第一步：根据偏好生成“草稿”分数
        draft_scores = {}
        total_weight = sum(item['weight'] for item in structure.values())
        
        if abs(total_weight - 1.0) > 0.01:
            raise ValueError(f"权重之和不为1 ({total_weight})")

        for name, config in structure.items():
            dist_type = config.get('type', 'normal')
            
            if dist_type == 'left_skewed':
                draft_scores[name] = self.dist_left_skewed(student_total_score)
            elif dist_type == 'right_skewed':
                draft_scores[name] = self.dist_right_skewed(student_total_score)
            elif dist_type == 'bimodal':
                draft_scores[name] = self.dist_bimodal()
            elif dist_type == 'discrete':
                levels = config.get('levels', [70, 80, 85, 90, 95])
                draft_scores[name] = self.dist_discrete(levels=levels)
            else: # normal
                draft_scores[name] = self.dist_normal(student_total_score)

        # 第二步：注入智能噪声 (调用上面的 apply_advanced_noise)
        # 如果 allowed_items 为空，默认所有科目都可注入（除非明确指定）
        allowed = noise_config.get('allowed_items', list(draft_scores.keys()))
        if not allowed: allowed = list(draft_scores.keys())

        draft_scores = self.apply_advanced_noise(
            draft_scores, 
            noise_ratio=noise_config.get('noise_ratio', 0.0),
            severity_mode=noise_config.get('severity_mode', 'random'),
            allowed_items=allowed
        )

        # 第三步：计算偏差并修正 (Target Alignment)
        current_weighted_sum = sum(draft_scores[k] * structure[k]['weight'] for k in draft_scores)
        diff = student_total_score - current_weighted_sum
        
        final_scores = {}
        for name, score in draft_scores.items():
            final_scores[name] = self._clamp(score + diff)
            
        # 第四步：锚点强行修正 (Anchor Fix)
        sorted_items = sorted(structure.items(), key=lambda x: x[1]['weight'], reverse=True)
        anchor_name = sorted_items[0][0]
        anchor_weight = sorted_items[0][1]['weight']
        
        post_clamp_total = sum(final_scores[k] * structure[k]['weight'] for k in final_scores)
        final_diff = student_total_score - post_clamp_total
        
        final_scores[anchor_name] += final_diff / anchor_weight
        final_scores[anchor_name] = self._clamp(final_scores[anchor_name])
        
        return {k: round(v, 1) for k, v in final_scores.items()}

    def _clamp(self, value):
        return max(0.0, min(100.0, value))