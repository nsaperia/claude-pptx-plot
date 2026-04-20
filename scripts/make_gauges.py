"""
Render three clean semicircle gauges in one PNG.

Why matplotlib and not shape-based?
  pptxgenjs addShape has no curve primitives, so shape-based gauges have
  to stair-step arcs. At slide scale that looks clunky (user feedback).
  Matplotlib renders smooth curves natively; insert as an image via
  pptxgenjs addImage. Consistent with the handoff's pattern for
  complex curved viz (make_quadrant.py, make_sankey.py).

Output: claude-pptx-plot/assets/gauges.png
"""

import os
import numpy as np
import matplotlib.pyplot as plt
import matplotlib.patches as mpatches
from matplotlib.patches import FancyArrowPatch

# Saperia palette
BG    = '#FFF5ED'
INK   = '#353745'
MUTED = '#8A7968'
RULE  = '#C9B9A8'
BERRY = '#B84C65'
STEEL = '#2D5F7C'
GOLD  = '#C5A55A'

TRACK_COLOR = '#E5DCCB'  # lighter version of RULE for the unfilled track

gauges = [
    dict(label='UTILIZATION',   value=67.6, domain=(55, 80),   good=72, fmt='{:.1f}%'),
    dict(label='WIN RATE',       value=46.9, domain=(40, 70),   good=55, fmt='{:.1f}%'),
    dict(label='EBITDA MARGIN',  value=26.3, domain=(10, 35),   good=28, fmt='{:.1f}%'),
]


def color_for_value(value, domain, good):
    """Red below 40% of range, gold in 40-75%, steel above."""
    lo, hi = domain
    frac = (value - lo) / (hi - lo)
    if frac < 0.40:
        return BERRY
    if frac < 0.75:
        return GOLD
    return STEEL


def draw_gauge(ax, g):
    """Draw one semicircle gauge into a matplotlib axis."""
    ax.set_aspect('equal')
    ax.set_xlim(-1.3, 1.3)
    ax.set_ylim(-0.8, 1.3)
    ax.axis('off')
    ax.set_facecolor(BG)

    # Geometry
    r_outer = 1.0
    r_inner = 0.72
    r_mid = (r_outer + r_inner) / 2
    track_width = (r_outer - r_inner)

    lo, hi = g['domain']
    frac = (g['value'] - lo) / (hi - lo)
    frac = max(0.0, min(1.0, frac))

    # In matplotlib Wedge, angles are degrees counterclockwise from 3 o'clock.
    # Semicircle: 180° (9 o'clock) to 0° (3 o'clock), going through 90° (top).
    # Fill arc is from 180° down to (180 - frac*180) — leftward = low, rightward = high.
    start_angle = 180
    end_filled  = 180 - frac * 180

    # Unfilled track (full half-circle in light gray)
    track = mpatches.Wedge(center=(0, 0), r=r_outer, theta1=0, theta2=180,
                            width=track_width, facecolor=TRACK_COLOR, edgecolor='none')
    ax.add_patch(track)

    # Filled portion in the value-appropriate color
    fill_color = color_for_value(g['value'], g['domain'], g.get('good', 0))
    filled = mpatches.Wedge(center=(0, 0), r=r_outer, theta1=end_filled, theta2=180,
                             width=track_width, facecolor=fill_color, edgecolor='none')
    ax.add_patch(filled)

    # Needle: from center out to the value angle on the midline of the track
    needle_angle_deg = 180 - frac * 180
    needle_angle_rad = np.deg2rad(needle_angle_deg)
    needle_len = r_mid + 0.05
    tip_x = needle_len * np.cos(needle_angle_rad)
    tip_y = needle_len * np.sin(needle_angle_rad)
    # Draw as a thin triangle for a sharper look
    needle = mpatches.FancyArrowPatch(
        (0, 0), (tip_x, tip_y),
        arrowstyle='-', mutation_scale=1, linewidth=2.5, color=INK, zorder=10
    )
    ax.add_patch(needle)
    # Hub
    hub = mpatches.Circle((0, 0), radius=0.09, facecolor=INK, edgecolor='none', zorder=11)
    ax.add_patch(hub)

    # Min / max tick labels at the ends of the arc
    tick_r = r_outer + 0.08
    ax.text(-tick_r, 0.02, str(lo),
            fontsize=9, fontfamily='DejaVu Sans', color=MUTED, ha='right', va='bottom')
    ax.text(tick_r, 0.02, str(hi),
            fontsize=9, fontfamily='DejaVu Sans', color=MUTED, ha='left', va='bottom')

    # Value (large) centered below the arc
    ax.text(0, -0.35, g['fmt'].format(g['value']),
            fontsize=26, fontfamily='serif', color=INK, ha='center', va='center')

    # Label (eyebrow style) below the value
    ax.text(0, -0.65, g['label'],
            fontsize=10, fontfamily='DejaVu Sans', color=MUTED,
            ha='center', va='center', fontweight='bold')


def main():
    n = len(gauges)
    fig, axes = plt.subplots(1, n, figsize=(12, 3.6), dpi=200)
    fig.patch.set_facecolor(BG)
    for ax, g in zip(axes, gauges):
        draw_gauge(ax, g)
    plt.subplots_adjust(left=0.02, right=0.98, top=0.96, bottom=0.04, wspace=0.05)

    script_dir = os.path.dirname(os.path.abspath(__file__))
    out_path = os.path.normpath(os.path.join(script_dir, '..', 'assets', 'gauges.png'))
    os.makedirs(os.path.dirname(out_path), exist_ok=True)
    plt.savefig(out_path, dpi=200, facecolor=BG, edgecolor='none', bbox_inches='tight', pad_inches=0.1)
    plt.close(fig)
    print(f'Wrote: {out_path}')


if __name__ == '__main__':
    main()
