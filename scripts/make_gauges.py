"""
Render ONE reusable arc PNG for gauge overlays.

Principle: minimize outsourcing. The PNG contains only what matplotlib
is uniquely good at — the three curved, colored arc zones. Everything
else on a gauge (tick marks, tick labels, target triangle, needle,
hub, all text, section headers, hairlines) is rendered natively in
pptxgenjs via drawGauge in src/plot.js.

The single PNG is reused for all gauges via pptxgenjs addImage:
- Standard orientation (higher = better): red-amber-green left to right
- Reversed orientation (lower = better): use flipH: true on addImage
- Neutral: reuse standard

Output: claude-pptx-plot/assets/gauge_arc.png
"""

import os
import numpy as np
import matplotlib.pyplot as plt
import matplotlib.patches as mpatches

# Saperia palette — muted zone tints (kept soft, not overstylized)
BG           = '#FFF5ED'
ZONE_DANGER  = '#E8D1D4'
ZONE_NEUTRAL = '#EFE5D3'
ZONE_GOOD    = '#D9E5D3'

def render_arc(out_path):
    # Figure sized so aspect ratio matches the arc's natural bounds:
    # semicircle on [-1, 1] × [0, 1] = 2:1
    fig = plt.figure(figsize=(4.0, 2.05), dpi=200)
    fig.patch.set_facecolor(BG)
    ax = fig.add_axes([0, 0, 1, 1])   # fill the figure exactly
    ax.set_aspect('equal')
    ax.set_xlim(-1.05, 1.05)
    ax.set_ylim(-0.05, 1.05)
    ax.axis('off')
    ax.set_facecolor(BG)

    r_outer = 1.0
    track_w = 0.22

    # Three 60° zones covering the upper half: 180° (9 o'clock) to 0° (3 o'clock)
    zones = [
        (120, 180, ZONE_DANGER),   # left third
        (60,  120, ZONE_NEUTRAL),  # middle third
        (0,    60, ZONE_GOOD),     # right third
    ]
    for theta1, theta2, color in zones:
        wedge = mpatches.Wedge(
            center=(0, 0), r=r_outer,
            theta1=theta1, theta2=theta2,
            width=track_w,
            facecolor=color, edgecolor='none',
        )
        ax.add_patch(wedge)

    plt.savefig(out_path, dpi=200, facecolor=BG, edgecolor='none',
                bbox_inches='tight', pad_inches=0.0)
    plt.close(fig)
    print(f'Wrote: {out_path}')


def main():
    script_dir = os.path.dirname(os.path.abspath(__file__))
    assets_dir = os.path.normpath(os.path.join(script_dir, '..', 'assets'))
    os.makedirs(assets_dir, exist_ok=True)

    # Clean up the old multi-gauge PNG if it exists
    old_path = os.path.join(assets_dir, 'gauges.png')
    if os.path.exists(old_path):
        os.remove(old_path)
        print(f'Removed old: {old_path}')

    render_arc(os.path.join(assets_dir, 'gauge_arc.png'))


if __name__ == '__main__':
    main()
