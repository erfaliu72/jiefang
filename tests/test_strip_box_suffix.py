"""strip_box_suffix / normalize_base_car_type 单元测试（20260807 方案一）。"""
import unittest
import sys
import os

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
# 避免触发 Flask 应用完整初始化：直接从 app 模块提取函数有副作用（创建 app），
# 因此这里独立实现相同逻辑做等价验证；如需严格一致可改用 import app。
from migrate_guidance_merge_base import strip_box_suffix, normalize_base_car_type


class TestStripBoxSuffix(unittest.TestCase):
    def test_multilevel_suffix(self):
        # 冷藏带尾板 → 先去带尾板再去冷藏
        self.assertEqual(strip_box_suffix('解放J6F锡柴150冷藏带尾板'), '解放J6F锡柴150')

    def test_chassis_suffix(self):
        self.assertEqual(strip_box_suffix('解放J6F锡柴150底盘'), '解放J6F锡柴150')
        self.assertEqual(strip_box_suffix('解放虎VR锡柴底盘'), '解放虎VR锡柴')

    def test_box_suffix(self):
        self.assertEqual(strip_box_suffix('解放J6F全柴190LNG厢货'), '解放J6F全柴190LNG')
        self.assertEqual(strip_box_suffix('解放J6F全柴190LNG高栏'), '解放J6F全柴190LNG')

    def test_joined_format(self):
        self.assertEqual(
            strip_box_suffix('纯电 / 虎六G中体 / 宁德时代 / 120度电 / 五档 / 厢货'),
            '纯电 / 虎六G中体 / 宁德时代 / 120度电 / 五档')

    def test_normal_car_type_unchanged(self):
        # 厢型词不在末尾 → 不变
        self.assertEqual(strip_box_suffix('解放轻卡4米2-虎6G 180混动-盟固利电池'), '解放轻卡4米2-虎6G 180混动-盟固利电池')
        self.assertEqual(strip_box_suffix('虎V载货国六'), '虎V载货国六')
        self.assertEqual(strip_box_suffix('虎VR单排平板-红色整车'), '虎VR单排平板-红色整车')
        self.assertEqual(strip_box_suffix('140宁德-单档（冷藏）'), '140宁德-单档（冷藏）')
        self.assertEqual(strip_box_suffix('测试车型-J6P'), '测试车型-J6P')

    def test_condition_prefix(self):
        # 去成色前缀
        self.assertEqual(normalize_base_car_type('新车解放J6F锡柴150冷藏'), '解放J6F锡柴150')
        self.assertEqual(normalize_base_car_type('二手车解放J6F全柴190LNG高栏'), '解放J6F全柴190LNG')

    def test_empty(self):
        self.assertEqual(strip_box_suffix(''), '')
        self.assertEqual(strip_box_suffix(None), '')


if __name__ == '__main__':
    unittest.main()
