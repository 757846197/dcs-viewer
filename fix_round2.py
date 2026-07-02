import os
os.chdir(r'D:\华为家庭存储\WorkBuddy\2026-06-03-09-46-04')
path = os.path.join('dcs_viewer', 'app.py')
with open(path, 'r', encoding='utf-8') as f:
    content = f.read()

# Fix 1: Plugging threshold 0.5 → 5.0
content = content.replace(
    'if not (prev_v < 0.5 and curr_v >= 0.5):',
    'if not (prev_v < 5.0 and curr_v >= 5.0):'
)
content = content.replace(
    'plug_on = plug_map.get(t_start_ts, 0) >= 0.5',
    'plug_on = plug_map.get(t_start_ts, 0) >= 5.0'
)
content = content.replace(
    'if plug_map.get(t_start_ts + offset, 0) >= 0.5:',
    'if plug_map.get(t_start_ts + offset, 0) >= 5.0:'
)
print('Fix 1: Plugging threshold 0.5 -> 5.0')

# Fix 2: Replace metric guide card with two side-by-side cards
old_guide = '''<div class="card" id="metricGuide" style="margin-bottom:16px">
            <div class="card-header" onclick="var b=this.parentElement.querySelector('.card-body');b.style.display=b.style.display==='none'?'':'none'" style="cursor:pointer">
                关键指标说明 <span style="font-size:11px;color:#94a3b8;font-weight:400;margin-left:8px">点击展开/折叠</span>
            </div>
            <div class="card-body" style="display:none;font-size:12px;line-height:1.8;color:#475569">
                <table style="font-size:11px"><thead><tr><th style="width:130px">指标</th><th style="width:360px">含义</th><th>适用</th></tr></thead><tbody>
                <tr><td style="font-weight:600">钻进位移</td><td>开口小车推进位置变化量（终点-起点），反映开口深度</td><td><span class="tag tag-info">开口</span></td></tr>
                <tr><td style="font-weight:600">推进压力峰值</td><td>推进进油压力在窗口内最大值（MPa），反映钻进阻力</td><td><span class="tag tag-info">开口</span></td></tr>
                <tr><td style="font-weight:600">转钎压力峰值</td><td>转钎进油压力在窗口内最大值（MPa），反映钻杆扭矩</td><td><span class="tag tag-info">开口</span></td></tr>
                <tr><td style="font-weight:600">钻透判定</td><td>推进位移骤增(>0.1m)且压力骤降(>20%)时判定已钻透</td><td><span class="tag tag-info">开口</span></td></tr>
                <tr><td style="font-weight:600">冲击状态</td><td>冲击进油压力是否激活（>0.5MPa），反映冲击锤状态</td><td><span class="tag tag-info">开口</span></td></tr>
                <tr><td style="font-weight:600">打泥量</td><td>打泥累计值，反映注入铁口的炮泥总量</td><td><span class="tag tag-warn">堵口</span></td></tr>
                <tr><td style="font-weight:600">打泥压力峰值</td><td>打泥压力窗口内最大值（MPa），19-21MPa为合理</td><td><span class="tag tag-warn">堵口</span></td></tr>
                <tr><td style="font-weight:600">保压时长</td><td>压力18-22MPa区间持续秒数，≥60s为合格保压</td><td><span class="tag tag-warn">堵口</span></td></tr>
                <tr><td style="font-weight:600">耗时</td><td>触发到结束完整时长（根据实际位移/压力动态判定）</td><td>全部</td></tr>
                </tbody></table>
            </div>
        </div>'''

new_guide = '''<div style="display:flex;gap:16px;margin-bottom:16px">
            <div class="card" style="flex:1;margin-bottom:0">
                <div class="card-header" onclick="var b=this.parentElement.querySelector('.card-body');b.style.display=b.style.display==='none'?'':'none'" style="cursor:pointer">
                    关键指标说明 <span style="font-size:11px;color:#94a3b8;font-weight:400;margin-left:8px">点击展开</span>
                </div>
                <div class="card-body" style="display:none;font-size:12px;line-height:1.8;color:#475569">
                    <table style="font-size:11px"><thead><tr><th style="width:110px">指标</th><th>含义</th></tr></thead><tbody>
                    <tr><td style="font-weight:600">钻进位移</td><td>开口小车推进位置变化量（终点-起点），反映开口深度</td></tr>
                    <tr><td style="font-weight:600">推进压力峰值</td><td>推进进油压力在窗口内最大值（MPa），反映钻进阻力</td></tr>
                    <tr><td style="font-weight:600">转钎压力峰值</td><td>转钎进油压力在窗口内最大值（MPa），反映钻杆扭矩</td></tr>
                    <tr><td style="font-weight:600">冲击状态</td><td>冲击进油压力是否激活（>0.5MPa），反映冲击锤状态</td></tr>
                    <tr><td style="font-weight:600">打泥量</td><td>打泥累计值，反映注入铁口的炮泥总量</td></tr>
                    <tr><td style="font-weight:600">打泥压力峰值</td><td>打泥压力窗口内最大值（MPa），19-21MPa为合理</td></tr>
                    <tr><td style="font-weight:600">耗时</td><td>触发到结束完整时长（根据实际位移/压力动态判定）</td></tr>
                    </tbody></table>
                </div>
            </div>
            <div class="card" style="flex:1;margin-bottom:0">
                <div class="card-header" onclick="var b=this.parentElement.querySelector('.card-body');b.style.display=b.style.display==='none'?'':'none'" style="cursor:pointer">
                    结果判定说明 <span style="font-size:11px;color:#94a3b8;font-weight:400;margin-left:8px">点击展开</span>
                </div>
                <div class="card-body" style="display:none;font-size:12px;line-height:1.8;color:#475569">
                    <table style="font-size:11px"><thead><tr><th style="width:70px">结果</th><th>判定条件</th></tr></thead><tbody>
                    <tr><td><span class="tag tag-ok">成功</span></td><td>开口：推进位移骤增(>0.1m)且压力骤降(>20%)；堵口：打泥量达标+保压≥60s</td></tr>
                    <tr><td><span class="tag tag-fail">失败</span></td><td>开口：无有效钻进位移；堵口：打泥量未达标</td></tr>
                    <tr><td><span class="tag tag-warn">未完成</span></td><td>开口：有钻进但未检测到钻透信号</td></tr>
                    <tr><td><span class="tag tag-info">未完整</span></td><td>堵口：打泥完成但保压不足</td></tr>
                    </tbody></table>
                </div>
            </div>
        </div>'''

if old_guide in content:
    content = content.replace(old_guide, new_guide)
    print('Fix 2: Metric guide split into 2 cards')
else:
    print('Fix 2: old_guide not found')

# Fix 3: Time format - remove year, use MM-DD HH:MM:SS (Beijing time)
# Change trigger_time display from substring(0,19) to custom format
old_trigger = 'html += \'<td>\'+(c.trigger_time||\'\').substring(0,19)+\'</td>\';'
new_trigger = '''html += '<td>'+formatBeijingTime(c.trigger_time)+'</td>';'''
content = content.replace(old_trigger, new_trigger)
print('Fix 3a: trigger_time format')

# Fix 3b: window_start/end format - also Beijing time MM-DD HH:MM:SS
old_win_start = '''var winStart = (c.window_start||'').substring(0,19);'''
new_win_start = '''var winStart = formatBeijingTime(c.window_start);'''
content = content.replace(old_win_start, new_win_start)

old_win_end = '''var winEnd = (c.window_end||'').substring(0,19);'''
new_win_end = '''var winEnd = formatBeijingTime(c.window_end);'''
content = content.replace(old_win_end, new_win_end)
print('Fix 3b: window time format')

# Add formatBeijingTime function to JS
old_js_func = 'function renderOpType(t){'
new_js_func = '''function formatBeijingTime(isoStr){
    if(!isoStr)return'--';
    var m=isoStr.match(/^(\\d{4})-(\\d{2})-(\\d{2})T(\\d{2}:\\d{2}:\\d{2})/);
    if(m)return m[2]+'-'+m[3]+' '+m[4];
    return isoStr.substring(0,19);
}
function renderOpType(t){'''
content = content.replace(old_js_func, new_js_func)
print('Fix 3c: added formatBeijingTime()')

# Fix 4: Stats cards - split opening and plugging stats separately
old_stats_js = '''function renderStats(){
    var openCycles = globalCycles.filter(function(c){return c.type=='opening';});
    var plugCycles = globalCycles.filter(function(c){return c.type=='plugging';});
    var openOk = openCycles.filter(function(c){return c.result=='success'||c.breakthrough;}).length;
    var plugOk = plugCycles.filter(function(c){return c.result=='success';}).length;

    var avgDur = 0;
    if(globalCycles.length>0){
        var total = globalCycles.reduce(function(s,c){return s+(c.duration_s||0);},0);
        avgDur = Math.round(total/globalCycles.length);
    }
    var durMin = Math.floor(avgDur/60);
    var durSec = avgDur%60;

    var html = '';
    html += '<div class="stat-card"><div class="stat-label">开口作业</div><div class="stat-value">'+openCycles.length+'</div><div class="stat-detail">'+openOk+' 次钻透成功</div></div>';
    html += '<div class="stat-card"><div class="stat-label">堵口作业</div><div class="stat-value">'+plugCycles.length+'</div><div class="stat-detail">'+plugOk+' 次完整完成</div></div>';
    html += '<div class="stat-card"><div class="stat-label">作业总数</div><div class="stat-value">'+globalCycles.length+'</div><div class="stat-detail">开口 + 堵口</div></div>';
    html += '<div class="stat-card"><div class="stat-label">平均耗时</div><div class="stat-value">'+durMin+'m'+durSec+'s</div><div class="stat-detail">每炉次</div></div>';'''

new_stats_js = '''function renderStats(){
    var openCycles = globalCycles.filter(function(c){return c.type=='opening';});
    var plugCycles = globalCycles.filter(function(c){return c.type=='plugging';});
    var openOk = openCycles.filter(function(c){return c.result=='success'||c.breakthrough;}).length;
    var plugOk = plugCycles.filter(function(c){return c.result=='success';}).length;

    var openDur = 0, plugDur = 0;
    if(openCycles.length>0){var t=openCycles.reduce(function(s,c){return s+(c.duration_s||0);},0);openDur=Math.round(t/openCycles.length);}
    if(plugCycles.length>0){var t=plugCycles.reduce(function(s,c){return s+(c.duration_s||0);},0);plugDur=Math.round(t/plugCycles.length);}

    var html = '';
    html += '<div class="stat-card"><div class="stat-label">开口作业</div><div class="stat-value">'+openCycles.length+' 次</div><div class="stat-detail">'+openOk+' 次钻透 | 均 '+Math.floor(openDur/60)+'m'+(openDur%60)+'s</div></div>';
    html += '<div class="stat-card"><div class="stat-label">堵口作业</div><div class="stat-value">'+plugCycles.length+' 次</div><div class="stat-detail">'+plugOk+' 次完成 | 均 '+Math.floor(plugDur/60)+'m'+(plugDur%60)+'s</div></div>';
    html += '<div class="stat-card"><div class="stat-label">开口成功率</div><div class="stat-value">'+(openCycles.length>0?Math.round(openOk/openCycles.length*100):0)+'%</div><div class="stat-detail">钻透/总数</div></div>';
    html += '<div class="stat-card"><div class="stat-label">堵口成功率</div><div class="stat-value">'+(plugCycles.length>0?Math.round(plugOk/plugCycles.length*100):0)+'%</div><div class="stat-detail">完成/总数</div></div>';'''

if old_stats_js in content:
    content = content.replace(old_stats_js, new_stats_js)
    print('Fix 4: Stats cards separated')
else:
    print('Fix 4: old_stats_js not found, trying alternative...')
    # Try with different whitespace
    idx = content.find('function renderStats(){')
    if idx > 0:
        print(f'Found renderStats at offset {idx}')
    else:
        print('renderStats not found!')

with open(path, 'w', encoding='utf-8') as f:
    f.write(content)
print('All fixes applied')
