import os
os.chdir(r'D:\华为家庭存储\WorkBuddy\2026-06-03-09-46-04')
path = os.path.join('dcs_viewer', 'app.py')
with open(path, 'r', encoding='utf-8') as f:
    content = f.read()

# Find _detect_opening_cycles function start
start = content.index('def _detect_opening_cycles(raw_data, sig):')
# Find _detect_plugging_cycles which follows
plug_start = content.index('def _detect_plugging_cycles(raw_data, sig):', start)

# New efficient opening detection
new_opening = '''def _detect_opening_cycles(raw_data, sig):
    """检测开口作业周期 — O(n) 高效版本"""
    pos = sorted(raw_data.get(sig["swing_pos"], []), key=lambda x: x[0])
    push_pos = sorted(raw_data.get(sig["push_pos"], []), key=lambda x: x[0])
    push_press = sorted(raw_data.get(sig["push_press"], []), key=lambda x: x[0])
    remote = sorted(raw_data.get(sig["remote"], []), key=lambda x: x[0])
    drill = sorted(raw_data.get(sig["drill_press"], []), key=lambda x: x[0])

    if len(pos) < 2 or len(remote) < 2:
        return []

    rem_map = {t.timestamp(): v for t, v in remote}
    cycles = []

    for i in range(1, len(pos)):
        prev_v, curr_v = pos[i - 1][1], pos[i][1]
        if not (prev_v < 90 and curr_v >= 90):
            continue
        t_cross = pos[i][0]
        t_cross_ts = t_cross.timestamp()

        remote_on = rem_map.get(t_cross_ts, 0) >= 0.5
        if not remote_on:
            for offset in (-1, 1):
                if rem_map.get(t_cross_ts + offset, 0) >= 0.5:
                    remote_on = True
                    break
        if not remote_on:
            continue

        t_start = t_cross
        t_end = t_start + timedelta(minutes=15)

        push_pos_change = 0.0
        push_press_peak = 0.0
        drill_press_peak = 0.0
        breakthrough_detected = False

        # Calculate push position change
        pos_before = [v for t, v in push_pos if t < t_start]
        pos_after = [(t, v) for t, v in push_pos if t_start <= t <= t_end]
        if pos_after:
            ref = pos_before[-1] if pos_before else pos_after[0][1]
            push_pos_change = pos_after[-1][1] - ref

        # Calculate pressure peaks and breakthrough detection
        if push_press:
            press_seg = [(t, v) for t, v in push_press if t_start <= t <= t_end]
            if press_seg:
                push_press_peak = max(v for _, v in press_seg)
                # Efficient breakthrough: sliding window of 3 consecutive samples
                pos_dict = {int(t.timestamp()): v for t, v in push_pos if t_start <= t <= t_end}
                for j in range(3, len(press_seg)):
                    p_curr = press_seg[j][1]
                    p_old = press_seg[j - 3][1]
                    t_curr_ts = int(press_seg[j][0].timestamp())
                    t_old_ts = int(press_seg[j - 3][0].timestamp())
                    dp_curr = pos_dict.get(t_curr_ts, pos_dict.get(t_curr_ts + 1, pos_dict.get(t_curr_ts - 1, 0)))
                    dp_old = pos_dict.get(t_old_ts, pos_dict.get(t_old_ts + 1, pos_dict.get(t_old_ts - 1, 0)))
                    delta_pos = dp_curr - dp_old
                    delta_press_ratio = (p_curr - p_old) / p_old if p_old > 0 else 0
                    if delta_pos > 0.1 and delta_press_ratio < -0.2:
                        breakthrough_detected = True
                        break

        if drill:
            drill_seg = [v for t, v in drill if t_start <= t <= t_end]
            if drill_seg:
                drill_press_peak = max(drill_seg)

        result = "success" if breakthrough_detected else ("incomplete" if push_pos_change > 0.01 else "fail")

        t_local_start = t_start + LOCAL_OFFSET
        t_local_end = t_end + LOCAL_OFFSET

        cycles.append({
            "machine": sig["name"],
            "type": "opening",
            "trigger_time": t_local_start.isoformat(),
            "window_start": t_local_start.isoformat(),
            "window_end": t_local_end.isoformat(),
            "duration_s": round((t_end - t_start).total_seconds(), 1),
            "push_pos_change": round(push_pos_change, 3),
            "push_press_peak": round(push_press_peak, 1),
            "drill_press_peak": round(drill_press_peak, 1),
            "breakthrough": breakthrough_detected,
            "result": result,
            "label": f"{sig['name']} 开口 {'成' if breakthrough_detected else '未'}钻透 {t_local_start.strftime('%H:%M:%S')} ~ {t_local_end.strftime('%H:%M:%S')}",
        })

    return cycles'''

# Replace in content
if 'for t, v in press_seg:' in content[start:plug_start]:
    content = content[:start] + new_opening + content[plug_start:]
    print('Replaced _detect_opening_cycles')
else:
    print('ERROR: pattern not found')

# Now also replace _detect_plugging_cycles
pl_start2 = content.index('def _detect_plugging_cycles(raw_data, sig):', start)
# Find next function or API route after it
next_api = content.index('@app.route("/api/analysis/cycles")', pl_start2)

new_plugging = '''def _detect_plugging_cycles(raw_data, sig):
    """检测堵口作业周期 — O(n) 高效版本"""
    cmd = sorted(raw_data.get(sig["mud_cmd"], []), key=lambda x: x[0])
    mud_press = sorted(raw_data.get(sig["mud_press"], []), key=lambda x: x[0])
    mud_qty = sorted(raw_data.get(sig["mud_qty"], []), key=lambda x: x[0])
    plug_select = sorted(raw_data.get(sig["plug_select"], []), key=lambda x: x[0])

    if len(cmd) < 2:
        return []

    plug_map = {t.timestamp(): v for t, v in plug_select}
    cycles = []

    for i in range(1, len(cmd)):
        prev_v, curr_v = cmd[i - 1][1], cmd[i][1]
        if not (prev_v < 0.5 and curr_v >= 0.5):
            continue
        t_start = cmd[i][0]
        t_start_ts = t_start.timestamp()

        plug_on = plug_map.get(t_start_ts, 0) >= 0.5
        if not plug_on:
            for offset in (-1, 1):
                if plug_map.get(t_start_ts + offset, 0) >= 0.5:
                    plug_on = True
                    break
        if not plug_on:
            continue

        t_end = t_start + timedelta(minutes=40)
        mud_press_peak = 0.0
        mud_qty_total = 0.0
        hold_duration_s = 0.0
        mud_fill_complete = False
        hold_complete = False

        press_seg = sorted([(t, v) for t, v in mud_press if t_start <= t <= t_end], key=lambda x: x[0])
        if press_seg:
            mud_press_peak = max(v for _, v in press_seg)
            hold_count = 0
            for _, v in press_seg:
                if 18 <= v <= 22:
                    hold_count += 1
                    if hold_count >= 60:
                        hold_duration_s = float(hold_count)
                        hold_complete = True
                        break
                else:
                    hold_count = 0

        qty_seg = sorted([(t, v) for t, v in mud_qty if t_start <= t <= t_end], key=lambda x: x[0])
        if qty_seg:
            mud_qty_total = qty_seg[-1][1]
            mud_fill_complete = mud_qty_total >= 10

        # Find retreat end
        for t, v in reversed(press_seg):
            if v < 5 and t > t_start + timedelta(seconds=30):
                t_end = t
                break

        result = "success" if (mud_fill_complete and hold_complete) else ("partial" if mud_fill_complete else "fail")

        t_local_start = t_start + LOCAL_OFFSET
        t_local_end = t_end + LOCAL_OFFSET

        cycles.append({
            "machine": sig["name"],
            "type": "plugging",
            "trigger_time": t_local_start.isoformat(),
            "window_start": t_local_start.isoformat(),
            "window_end": t_local_end.isoformat(),
            "duration_s": round((t_end - t_start).total_seconds(), 1),
            "mud_press_peak": round(mud_press_peak, 1),
            "mud_qty": round(mud_qty_total, 1),
            "hold_duration_s": round(hold_duration_s, 0),
            "mud_filled": mud_fill_complete,
            "hold_ok": hold_complete,
            "result": result,
            "label": f"{sig['name']} 堵口 {'完' if mud_fill_complete else '未完'} {t_local_start.strftime('%H:%M:%S')} ~ {t_local_end.strftime('%H:%M:%S')}",
        })

    return cycles'''

content = content[:pl_start2] + new_plugging + content[next_api:]
print('Replaced _detect_plugging_cycles')

with open(path, 'w', encoding='utf-8') as f:
    f.write(content)
print('Done')
