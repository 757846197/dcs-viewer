"""
动态钻透判定与铁口深度计算模块

原理:
1. 每次开口/堵口周期开始时，自动采集小车初始位置作为动态基线
2. 周期内采集行程数据，计算相对位移 = (当前值 - 基线) × 斜率修正 + 偏移补偿
3. 钻透判定综合三项指标: 位移趋势、速度拐点、偏移修正后的绝对值
4. 铁口深度 = max(相对位移) + 偏移基线补偿量

核心优势: 不依赖固定阈值，自动适应编码器未标定/偏移场景
"""

import logging
from datetime import datetime, timedelta
from collections import deque

log = logging.getLogger(__name__)

class DynamicBreakthroughDetector:
    """动态钻透检测器 — 基于行程的实时判定

    使用方法:
        detector = DynamicBreakthroughDetector(calib)
        detector.feed(timestamp, position_value)
        result = detector.check_breakthrough()
    """

    def __init__(self, calibration=None):
        self.calib = calibration or {}
        self.offset_baseline = self.calib.get("offset_baseline", 0.0)
        self.slope_correction = self.calib.get("slope_correction", 1.0)
        self.travel_max = self.calib.get("travel_range_max", 3.0)

        # 基线窗口
        self.baseline_window = deque(maxlen=10)
        self.baseline = None  # 动态基线
        self.baseline_established = False

        # 运动追踪
        self.positions = []  # [(ts, raw_val, corrected_val)]
        self.max_corrected = 0.0

        # 钻透判定状态
        self.breakthrough_detected = False
        self.breakthrough_time = None
        self.breakthrough_position = 0.0

        # 速度分析用
        self.velocities = deque(maxlen=20)
        self._prev_pos = None
        self._prev_ts = None

    def feed(self, ts, raw_value):
        """喂入一个数据点"""
        corrected = self._correct(raw_value)
        self.positions.append((ts, raw_value, corrected))

        if corrected > self.max_corrected:
            self.max_corrected = corrected

        # 基线采集
        if not self.baseline_established:
            self.baseline_window.append(corrected)
            if len(self.baseline_window) >= 5:
                self.baseline = self._compute_baseline()
                self.baseline_established = True
            return

        # 速度计算
        if self._prev_pos is not None and self._prev_ts is not None:
            dt = (ts - self._prev_ts).total_seconds()
            if dt > 0:
                vel = (corrected - self._prev_pos) / dt
                self.velocities.append((ts, vel))
        self._prev_pos = corrected
        self._prev_ts = ts

        # 钻透检测
        if not self.breakthrough_detected and self.baseline_established:
            if self._evaluate_breakthrough(ts):
                self.breakthrough_detected = True
                self.breakthrough_time = ts
                self.breakthrough_position = corrected

    def _correct(self, raw):
        """编码器偏移校正"""
        return (raw + self.offset_baseline) * self.slope_correction

    def _compute_baseline(self):
        """IQR 去噪后取中位数作为基线"""
        vals = sorted(self.baseline_window)
        n = len(vals)
        q1 = vals[n // 4]
        q3 = vals[3 * n // 4]
        iqr = q3 - q1
        lo, hi = q1 - 1.5 * iqr, q3 + 1.5 * iqr
        filtered = [v for v in vals if lo <= v <= hi]
        if not filtered:
            filtered = vals
        return filtered[len(filtered) // 2]

    def _evaluate_breakthrough(self, ts):
        """多维钻透判定（不依赖固定阈值）"""
        if self.baseline is None or len(self.positions) < 10:
            return False

        travel = self.max_corrected - self.baseline
        score = 0

        # 指标1: 相对行程占比（是否达到行程范围的60%以上）
        travel_ratio = travel / self.travel_max if self.travel_max > 0 else 0
        if travel_ratio >= 0.6:
            score += 2
        elif travel_ratio >= 0.4:
            score += 1

        # 指标2: 速度拐点（快速推进后减速 → 钻透）
        if len(self.velocities) >= 10:
            recent_vels = [v for _, v in list(self.velocities)[-10:]]
            max_vel = max(recent_vels)
            last_vel = recent_vels[-1] if recent_vels else 0
            if max_vel > 0.01:
                slowdown_ratio = last_vel / max_vel
                if slowdown_ratio < 0.3:
                    score += 2  # 明显减速
                elif slowdown_ratio < 0.5:
                    score += 1

        # 指标3: 位移稳定（最后几个点变化小，说明停止）
        if len(self.positions) >= 5:
            last_corrected = [c for _, _, c in self.positions[-5:]]
            variation = max(last_corrected) - min(last_corrected)
            if variation < 0.02:
                score += 1

        # 综合判定: score >= 4 即为钻透
        return score >= 4

    def get_depth(self):
        """计算修正后铁口深度"""
        if not self.baseline_established or self.baseline is None:
            return None
        depth_mm = int((self.max_corrected - self.baseline) * 1000)
        return {
            "depth_mm": depth_mm,
            "depth_m": round(depth_mm / 1000, 3),
            "baseline_m": round(self.baseline, 3),
            "max_position_m": round(self.max_corrected, 3),
            "breakthrough_m": round(self.breakthrough_position, 3) if self.breakthrough_detected else None,
            "breakthrough_detected": self.breakthrough_detected,
            "breakthrough_time": self.breakthrough_time.isoformat() if self.breakthrough_time else None,
            "travel_ratio": round((self.max_corrected - self.baseline) / self.travel_max, 3) if self.travel_max > 0 else 0,
        }


def run_dynamic_analysis(start_time, end_time, cycle_type, position_signal,
                         calibration, machine=""):
    """完整动态分析流程 — 供 API 调用

    Args:
        start_time, end_time: UTC 时间字符串 "2026-07-04T00:00:00Z"
        cycle_type: "opening" | "plugging"
        position_signal: 小车位置信号名
        calibration: 编码器校准 dict
        machine: 设备名

    Returns:
        dict: {breakthrough, depth, metrics}
    """
    from dcs_platform.core.influx_client import fetch_timeseries

    try:
        data = fetch_timeseries(start_time, end_time, [position_signal], timeout_ms=60000)
        pts = data.get(position_signal, [])
    except Exception as e:
        log.warning("fetch_timeseries failed: %s", e)
        pts = []

    if not pts:
        return {
            "error": "no_data",
            "breakthrough_detected": False,
            "depth": None
        }

    detector = DynamicBreakthroughDetector(calibration)
    for ts, val in sorted(pts, key=lambda x: x[0]):
        detector.feed(ts, float(val))

    depth_result = detector.get_depth()

    return {
        "machine": machine,
        "cycle_type": cycle_type,
        "breakthrough_detected": detector.breakthrough_detected,
        "breakthrough_time": detector.breakthrough_time.isoformat() if detector.breakthrough_time else None,
        "breakthrough_position_m": round(detector.breakthrough_position, 3) if detector.breakthrough_detected else None,
        "depth_mm": depth_result["depth_mm"] if depth_result else None,
        "depth_m": depth_result["depth_m"] if depth_result else None,
        "baseline_m": depth_result["baseline_m"] if depth_result else None,
        "max_position_m": depth_result["max_position_m"] if depth_result else None,
        "travel_ratio": depth_result["travel_ratio"] if depth_result else 0,
        "total_points": len(pts),
    }
