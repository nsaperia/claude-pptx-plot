"""
Render six KPI gauges in a 2 × 3 grid with two section headers.

Structure emulates a reference the user liked:
  Section 1 "THE MONEY WORKS"     — steel blue eyebrow
    three gauges, targets above
  Section 2 "THE PEOPLE PROBLEM"  — berry eyebrow
    three gauges, targets below

Each gauge:
  title (bold sans) + subtitle (italic muted)
  arc with three soft-tinted zones (danger / middle / good)
  tick labels on the arc
  triangle marker indicating target
  needle pointing at current value + hub
  center value (large serif) + "Target: X" + above/below delta

Output: claude-pptx-plot/assets/gauges.png
"""

import os
import numpy as np
import matplotlib.pyplot as plt
import matplotlib.patches as mpatches

# Saperia palette
BG        = '#FFF5ED'
BG_RAISED = '#FBEFE4'
INK       = '#353745'
MUTED     = '#8A7968'
RULE      = '#C9B9A8'
BERRY     = '#B84C65'
STEEL     = '#2D5F7C'
GOLD      = '#C5A55A'

# Muted zone tints (soft, not overstylized)
ZONE_DANGER  = '#E8D1D4'  # soft berry
ZONE_NEUTRAL = '#EFE5D3'  # soft beige
ZONE_GOOD    = '#D9E5D3'  # soft sage

# KPI definitions. `higher_is_better` flips zone order for metrics where low is good.
gauges = [
    # Row 1: THE MONEY WORKS
    dict(section=0, title='Realization Rate',   subtitle='Billed vs. worked value',
         domain=(50, 100), value=93.0, target=90.0, higher_is_better=True,
         value_fmt='{:.0f}%', delta_fmt='{:+.1f} above target',
         ticks=[60, 70, 80, 90]),
    dict(section=0, title='Collection Rate',    subtitle='Collected vs. billed',
         domain=(50, 100), value=93.0, target=90.0, higher_is_better=True,
         value_fmt='{:.0f}%', delta_fmt='{:+.1f} above target',
         ticks=[60, 70, 80, 90]),
    dict(section=0, title='EBITDA Margin',      subtitle='Profitability after expenses',
         domain=(0, 60), value=26.3, target=25.0, higher_is_better=True,
         value_fmt='{:.1f}%', delta_fmt='{:+.1f} above target',
         ticks=[12, 24, 36, 48]),
    # Row 2: THE PEOPLE PROBLEM
    dict(section=1, title='Utilization Rate',   subtitle='Billable vs. total hours',
         domain=(30, 90), value=67.6, target=72.0, higher_is_better=True,
         value_fmt='{:.0f}%', delta_fmt='{:+.1f} below target',
         ticks=[42, 54, 66, 78]),
    dict(section=1, title='Annualized Turnover', subtitle='Departures / headcount',
         domain=(0, 40), value=7.0, target=15.0, higher_is_better=False,
         value_fmt='{:.0f}%', delta_fmt='{:+.1f} below target',
         ticks=[8, 16, 24, 32]),
    dict(section=1, title='Leverage Ratio',     subtitle='Staff per equity partner',
         domain=(0, 15), value=10.0, target=6.0, higher_is_better=None,
         value_fmt='{:.1f}×', delta_fmt='{:+.1f} above target',
         ticks=[3, 6, 9, 12]),
]


def zone_colors(higher_is_better):
    """Return (left_tint, mid_tint, right_tint) along the arc from 180° to 0°."""
    if higher_is_better is True:
        return ZONE_DANGER, ZONE_NEUTRAL, ZONE_GOOD
    if higher_is_better is False:
        return ZONE_GOOD, ZONE_NEUTRAL, ZONE_DANGER
    # None: neutral reference (no good/bad zones, e.g. ratio with target marker)
    return ZONE_NEUTRAL, ZONE_NEUTRAL, ZONE_NEUTRAL


def needle_color(status_good):
    return STEEL if status_good else BERRY


def status_is_good(g):
    """Compare value to target given higher-is-better orientation."""
    if g['higher_is_better'] is True:
        return g['value'] >= g['target']
    if g['higher_is_better'] is False:
        return g['value'] <= g['target']
    # None: use absolute tolerance
    return abs(g['value'] - g['target']) / max(1, g['target']) < 0.25


def draw_gauge(ax, g):
    ax.set_aspect('equal')
    ax.set_xlim(-1.55, 1.55)
    ax.set_ylim(-1.05, 1.65)
    ax.axis('off')
    ax.set_facecolor(BG)

    lo, hi = g['domain']
    r_outer = 0.9       # slightly smaller so tick labels don't crowd
    track_w = 0.17
    r_inner = r_outer - track_w

    # Colored zones — three segments of 60° each
    z1, z2, z3 = zone_colors(g['higher_is_better'])
    for theta1, theta2, color in [(120, 180, z1), (60, 120, z2), (0, 60, z3)]:
        wedge = mpatches.Wedge(center=(0, 0), r=r_outer,
                                theta1=theta1, theta2=theta2,
                                width=track_w, facecolor=color, edgecolor='none')
        ax.add_patch(wedge)

    # Tick marks AT tick positions + numeric label above the arc
    for tv in g['ticks']:
        frac = (tv - lo) / (hi - lo)
        angle_deg = 180 - frac * 180
        angle_rad = np.deg2rad(angle_deg)
        # Short tick outside the arc
        x1 = np.cos(angle_rad) * (r_outer + 0.015)
        y1 = np.sin(angle_rad) * (r_outer + 0.015)
        x2 = np.cos(angle_rad) * (r_outer + 0.07)
        y2 = np.sin(angle_rad) * (r_outer + 0.07)
        ax.plot([x1, x2], [y1, y2], color=MUTED, linewidth=0.8)
        # Label above the tick
        lx = np.cos(angle_rad) * (r_outer + 0.17)
        ly = np.sin(angle_rad) * (r_outer + 0.17)
        ax.text(lx, ly, str(tv), fontsize=8, color=MUTED,
                fontfamily='DejaVu Sans',
                ha='center', va='center')

    # Min / max labels at the extremes (anchored slightly below axis line)
    ax.text(-r_outer - 0.05, -0.05, str(lo), fontsize=8, color=MUTED,
            fontfamily='DejaVu Sans', ha='right', va='top')
    ax.text(r_outer + 0.05, -0.05, str(hi), fontsize=8, color=MUTED,
            fontfamily='DejaVu Sans', ha='left', va='top')

    # Target marker — small triangle just above the tick labels, with its
    # apex pointing at the arc. Tick labels sit at r_outer + 0.17, so
    # triangle base at r_outer + 0.24 clears them.
    tfrac = (g['target'] - lo) / (hi - lo)
    tfrac = max(0.0, min(1.0, tfrac))
    t_angle = 180 - tfrac * 180
    trad = np.deg2rad(t_angle)
    tx = np.cos(trad) * (r_outer + 0.24)
    ty = np.sin(trad) * (r_outer + 0.24)
    tri = mpatches.Polygon(
        [(tx - 0.04, ty + 0.06),
         (tx + 0.04, ty + 0.06),
         (tx,        ty)],
        facecolor=STEEL, edgecolor='none')
    ax.add_patch(tri)

    # Needle — from hub to current value position on the mid-line of the track
    vfrac = (g['value'] - lo) / (hi - lo)
    vfrac = max(0.0, min(1.0, vfrac))
    v_angle = 180 - vfrac * 180
    vrad = np.deg2rad(v_angle)
    r_mid = (r_outer + r_inner) / 2
    tip_x = np.cos(vrad) * (r_mid + 0.03)
    tip_y = np.sin(vrad) * (r_mid + 0.03)
    ncolor = BERRY if g['higher_is_better'] in (True, False) and not status_is_good(g) else BERRY
    # Keep needle berry for contrast against tinted zones, per reference image
    ax.plot([0, tip_x], [0, tip_y], color=BERRY, linewidth=2.2, solid_capstyle='round', zorder=10)
    hub = mpatches.Circle((0, 0), radius=0.08, facecolor=INK, edgecolor='none', zorder=11)
    ax.add_patch(hub)
    hub_dot = mpatches.Circle((0, 0), radius=0.025, facecolor=BERRY, edgecolor='none', zorder=12)
    ax.add_patch(hub_dot)

    # Center value (large serif), target, delta
    ax.text(0, -0.30, g['value_fmt'].format(g['value']),
            fontsize=24, fontfamily='serif', color=INK,
            ha='center', va='top')
    target_text = 'Target: <{:.0f}%'.format(g['target']) if g['higher_is_better'] is False else 'Target: {:g}{}'.format(
        g['target'], '%' if '%' in g['value_fmt'] else ('×' if '×' in g['value_fmt'] else '')
    )
    ax.text(0, -0.68, target_text,
            fontsize=9, fontfamily='serif', fontstyle='italic', color=INK,
            ha='center', va='top')
    # Delta status
    good = status_is_good(g)
    delta = g['value'] - g['target']
    delta_color = STEEL if good else BERRY
    # Format the delta string
    if g['higher_is_better'] is False:
        # Below target is "good" — phrase as points below
        below_points = g['target'] - g['value']
        delta_str = '{:+.1f} below target'.format(-below_points)
    else:
        delta_str = '{:+.1f} {} target'.format(delta, 'above' if delta >= 0 else 'below')
    ax.text(0, -0.90, delta_str,
            fontsize=9, color=delta_color,
            fontfamily='DejaVu Sans', fontweight='bold',
            ha='center', va='top')


def main():
    # Figure: 12 wide × 8.0 tall at 200 dpi (taller to give each gauge
    # vertical breathing room — title + subtitle + arc + value stack)
    fig = plt.figure(figsize=(12, 8.0), dpi=200)
    fig.patch.set_facecolor(BG)

    # 2 section headers + 2 rows of 3 gauges. Use GridSpec for control.
    gs = fig.add_gridspec(
        nrows=4, ncols=3,
        height_ratios=[0.12, 1.0, 0.12, 1.0],
        hspace=0.12, wspace=0.0,
        left=0.03, right=0.97, top=0.97, bottom=0.03,
    )

    # Section headers
    def header(ax, text, color):
        ax.axis('off')
        ax.set_xlim(0, 1); ax.set_ylim(0, 1)
        ax.text(0.02, 0.3, text, fontsize=11, color=color,
                fontfamily='DejaVu Sans', fontweight='bold',
                ha='left', va='center')
        # Thin rule under the label
        ax.plot([0.02, 0.97], [0.05, 0.05], color=RULE, linewidth=0.5)

    header_axes = [
        (0, 'THE MONEY WORKS',   STEEL),
        (2, 'THE PEOPLE PROBLEM', BERRY),
    ]
    for row, text, col in header_axes:
        # Header spans all three columns
        ax = fig.add_subplot(gs[row, :])
        header(ax, text, col)

    # Gauge grid
    for i, g in enumerate(gauges):
        row_block = 1 if g['section'] == 0 else 3
        col = i % 3
        ax = fig.add_subplot(gs[row_block, col])

        # Per-gauge title + subtitle, positioned above the gauge axis.
        # ylim is extended to 1.65 so subtitle + title sit above the arc
        # (which tops out around y≈1.07 after the tick-label offset).
        ax.text(0, 1.55, g['title'], fontsize=12, color=INK,
                fontfamily='DejaVu Sans', fontweight='bold',
                ha='center', va='center')
        ax.text(0, 1.32, g['subtitle'], fontsize=9, color=MUTED,
                fontfamily='serif', fontstyle='italic',
                ha='center', va='center')

        draw_gauge(ax, g)

    script_dir = os.path.dirname(os.path.abspath(__file__))
    out_path = os.path.normpath(os.path.join(script_dir, '..', 'assets', 'gauges.png'))
    os.makedirs(os.path.dirname(out_path), exist_ok=True)
    plt.savefig(out_path, dpi=200, facecolor=BG, edgecolor='none',
                bbox_inches='tight', pad_inches=0.15)
    plt.close(fig)
    print(f'Wrote: {out_path}')


if __name__ == '__main__':
    main()
