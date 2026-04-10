import matplotlib
matplotlib.use('Agg')
import matplotlib.pyplot as plt
import matplotlib.patches as mpatches
import numpy as np

# ===== データ =====
daily_data = [
    ('1/25','日',39400,10), ('1/26','月',20500,5),  ('1/27','火',28600,5),
    ('1/28','水',3500,1),   ('1/29','木',11600,4),  ('1/30','金',18000,6),
    ('1/31','土',17300,6),  ('2/1','日',31600,11),  ('2/2','月',26100,6),
    ('2/3','火',47200,12),  ('2/4','水',31199,9),   ('2/5','木',47800,12),
    ('2/6','金',40000,11),  ('2/7','土',20500,5),   ('2/8','日',17000,5),
    ('2/9','月',13600,3),   ('2/10','火',0,0),       ('2/11','水',17500,5),
    ('2/12','木',28700,9),  ('2/13','金',27400,9),  ('2/14','土',72200,19),
    ('2/15','日',37000,11), ('2/16','月',47300,12), ('2/17','火',42700,13),
    ('2/18','水',28900,9),  ('2/19','木',10700,2),  ('2/20','金',35600,10),
    ('2/21','土',54700,11), ('2/22','日',62900,18), ('2/23','月',33400,9),
    ('2/24','火',52700,13), ('2/25','水',11750,4),  ('2/26','木',12150,4),
    ('2/27','金',32250,10), ('2/28','土',22550,7),  ('3/1','日',13600,5),
    ('3/2','月',42300,10),  ('3/3','火',30700,9),   ('3/4','水',26300,9),
    ('3/5','木',18500,5),   ('3/6','金',62000,12),  ('3/7','土',72100,17),
    ('3/8','日',0,0),        ('3/9','月',35850,8),   ('3/10','火',33300,9),
    ('3/11','水',35000,10), ('3/12','木',67900,17), ('3/13','金',53000,11),
    ('3/14','土',27400,8),  ('3/15','日',31200,9),  ('3/16','月',31850,7),
    ('3/17','火',41800,11), ('3/18','水',68100,15), ('3/19','木',63900,15),
    ('3/20','金',79050,16), ('3/21','土',40450,13), ('3/22','日',43000,12),
    ('3/23','月',104600,15),('3/24','火',55300,17), ('3/25','水',21800,7),
    ('3/26','木',47500,15), ('3/27','金',108700,25),('3/28','土',34600,10),
    ('3/29','日',25200,7),  ('3/30','月',6800,2),   ('3/31','火',34800,9),
]

dates   = [d[0] for d in daily_data]
days    = [d[1] for d in daily_data]
sales   = [d[2] for d in daily_data]
customers = [d[3] for d in daily_data]

def day_color(day):
    if day == '金': return '#C9A96E'
    if day in ('土','日'): return '#E87E7E'
    return '#4E8EF7'

colors = [day_color(d) for d in days]

# 7日移動平均
ma7 = []
for i in range(len(sales)):
    if i < 6:
        ma7.append(None)
    else:
        ma7.append(sum(sales[i-6:i+1]) / 7)

# ===== スタイル =====
plt.rcParams.update({
    'font.family': ['IPAGothic', 'Noto Sans CJK JP', 'DejaVu Sans'],
    'axes.facecolor': '#16213E',
    'figure.facecolor': '#1A1A2E',
    'axes.edgecolor': '#0F3460',
    'axes.labelcolor': '#AAAAAA',
    'xtick.color': '#888888',
    'ytick.color': '#888888',
    'grid.color': '#1E3A5F',
    'grid.linestyle': '--',
    'grid.alpha': 0.5,
    'text.color': '#EEEEEE',
})

fig, axes = plt.subplots(3, 1, figsize=(20, 18))
fig.patch.set_facecolor('#1A1A2E')
fig.suptitle('シーシャバー 売上データ分析  2026.1/25–3/31',
             fontsize=16, color='#C9A96E', y=0.98, fontweight='bold')

x = np.arange(len(dates))

# ===== グラフ① 日別売上 =====
ax1 = axes[0]
ax1.set_facecolor('#16213E')
bars = ax1.bar(x, sales, color=colors, width=0.7, alpha=0.85, zorder=2)
ma_x = [i for i, v in enumerate(ma7) if v is not None]
ma_y = [v for v in ma7 if v is not None]
ax1.plot(ma_x, ma_y, color='#00D4AA', linewidth=2, label='7日移動平均', zorder=3)
ax1.set_title('日別売上推移', color='#C9A96E', fontsize=13, pad=10)
ax1.set_ylabel('売上（円）', fontsize=10)
ax1.yaxis.set_major_formatter(plt.FuncFormatter(lambda v, _: f'¥{int(v/1000)}k'))
ax1.set_xticks(x[::2])
ax1.set_xticklabels([f"{dates[i]}\n{days[i]}" for i in x[::2]], fontsize=7.5)
ax1.grid(axis='y', zorder=0)
ax1.axhline(y=50000, color='#ffffff', alpha=0.15, linestyle=':', linewidth=1)
ax1.set_xlim(-0.8, len(x) - 0.2)

patch_fri = mpatches.Patch(color='#C9A96E', label='金曜')
patch_wknd = mpatches.Patch(color='#E87E7E', label='土日')
patch_wkd  = mpatches.Patch(color='#4E8EF7', label='平日')
patch_ma   = mpatches.Patch(color='#00D4AA', label='7日移動平均')
ax1.legend(handles=[patch_fri, patch_wknd, patch_wkd, patch_ma],
           facecolor='#0F3460', edgecolor='#0F3460', fontsize=9,
           loc='upper left', labelcolor='white')

# ===== グラフ② 月別売上 =====
ax2 = axes[1]
ax2.set_facecolor('#16213E')

month_labels = ['1月\n（7日間）', '2月\n（28日間）', '3月\n（31日間）']
month_sales  = [138900, 905399, 1356600]
month_avg    = [138900/7, 905399/28, 1356600/31]
month_colors = ['#C9A96E', '#4E8EF7', '#00D4AA']

xm = np.arange(3)
bw = 0.35
b1 = ax2.bar(xm - bw/2, month_sales, width=bw, color=month_colors, alpha=0.8, label='月間売上', zorder=2)
ax2_r = ax2.twinx()
ax2_r.set_facecolor('#16213E')
b2 = ax2_r.bar(xm + bw/2, month_avg, width=bw, color=month_colors, alpha=0.45, label='1日平均', zorder=2)

for bar, val in zip(b1, month_sales):
    ax2.text(bar.get_x() + bar.get_width()/2, bar.get_height() + 15000,
             f'¥{val:,}', ha='center', va='bottom', fontsize=10, color='#EEE')
for bar, val in zip(b2, month_avg):
    ax2_r.text(bar.get_x() + bar.get_width()/2, bar.get_height() + 800,
               f'¥{int(val):,}', ha='center', va='bottom', fontsize=9, color='#AAA')

ax2.set_title('月別売上合計 ／ 1日平均', color='#C9A96E', fontsize=13, pad=10)
ax2.set_xticks(xm)
ax2.set_xticklabels(month_labels, fontsize=11)
ax2.set_ylabel('月間売上（円）', fontsize=10)
ax2_r.set_ylabel('1日平均（円）', fontsize=10, color='#AAA')
ax2.yaxis.set_major_formatter(plt.FuncFormatter(lambda v, _: f'¥{int(v/10000)}万'))
ax2_r.yaxis.set_major_formatter(plt.FuncFormatter(lambda v, _: f'¥{int(v/1000)}k'))
ax2.grid(axis='y', zorder=0)
ax2_r.tick_params(colors='#888')

# ===== グラフ③ 曜日別平均 =====
ax3 = axes[2]
ax3.set_facecolor('#16213E')

dow_labels = ['月', '火', '水', '木', '金', '土', '日']
dow_avg    = [36230, 36710, 27116, 34306, 50667, 40200, 30090]
dow_colors = ['#4E8EF7','#4E8EF7','#4E8EF7','#4E8EF7','#C9A96E','#E87E7E','#E87E7E']

xd = np.arange(7)
b3 = ax3.bar(xd, dow_avg, color=dow_colors, width=0.6, alpha=0.85, zorder=2)
for bar, val in zip(b3, dow_avg):
    ax3.text(bar.get_x() + bar.get_width()/2, bar.get_height() + 400,
             f'¥{val:,}', ha='center', va='bottom', fontsize=10, color='#EEE')

ax3.axhline(y=sum(dow_avg)/7, color='#00D4AA', linestyle='--', linewidth=1.5,
            label=f'全曜日平均 ¥{int(sum(dow_avg)/7):,}', zorder=3)
ax3.set_title('曜日別 1日平均売上', color='#C9A96E', fontsize=13, pad=10)
ax3.set_xticks(xd)
ax3.set_xticklabels(dow_labels, fontsize=14)
ax3.set_ylabel('平均売上（円）', fontsize=10)
ax3.yaxis.set_major_formatter(plt.FuncFormatter(lambda v, _: f'¥{int(v/1000)}k'))
ax3.grid(axis='y', zorder=0)
ax3.set_ylim(0, max(dow_avg) * 1.2)
ax3.legend(facecolor='#0F3460', edgecolor='#0F3460', fontsize=9, labelcolor='white')

plt.tight_layout(rect=[0, 0, 1, 0.97])
plt.savefig('/home/user/kuon_shisha/sales_chart.png', dpi=150, bbox_inches='tight',
            facecolor='#1A1A2E')
print("saved")
