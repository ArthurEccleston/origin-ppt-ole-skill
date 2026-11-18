r"""Origin / OPJU reusable operations.

Tested with OriginPro 2025b, originpro 1.1.15, OriginExt 1.2.6, Python 3.13.
Windows only: originpro talks to a local Origin instance through COM.

Typical use from Claude Code:
    import origin_ops as oo
    oo.connect_standalone()
    gp = oo.open_project(r"...\graph.opju")
    data = oo.read_all_plots(gp)
    oo.set_axis_kelvin(gp[0], 'x')
    oo.save_export(gp, r"...\graph_K.opju", r"...\graph_K.png")
    oo.exit()

CLI:
    python origin_ops.py read --opju graph.opju --json data.json
    python origin_ops.py axis --opju graph.opju --axis x --out graph_K.opju
    python origin_ops.py fit  --opju graph_K.opju --plot 1 --model logistic --out graph_fit.opju
"""

import argparse
import json
import math
import os
import sys

import originpro as op

try:
    import numpy as np
    from scipy.optimize import curve_fit
except Exception:  # keep read/axis usable even without scipy
    np = None
    curve_fit = None


# ---------------------------------------------------------------------------
# connection / project
# ---------------------------------------------------------------------------

def connect_standalone(show=False):
    """Create/connect to an Origin instance for a standalone OPJU."""
    op.set_show(show)
    return op


def connect_attached(show=False):
    """Attach to an already-running Origin OLE server.

    Use this after activating an embedded OLE object in PowerPoint/WPS when
    you need to edit the embedded project in place.
    """
    op.attach()
    try:
        op.set_show(show)
    except Exception:
        pass
    return op


def open_project(path):
    """Open an OPJU and return graph pages.

    Warning: op.open() sometimes returns False for older projects even though
    the graph loads. Always check graph_list(), not the return value.
    """
    op.open(path)
    return op.graph_list()


def exit_origin():
    try:
        op.exit()
    except Exception:
        pass


# ---------------------------------------------------------------------------
# axis / title conversion
# ---------------------------------------------------------------------------

def to_kelvin_title(t):
    """Convert Origin rich-text Celsius axis title to Kelvin.

    Handles Origin rich text such as:
      \\p120(Temperature \\(40))\\g(\\p120(°))\\p120(C\\(41))
      \\p109(Temperature \\(40))\\f:Times New Roman(\\p109(℃))\\p109(\\(41))
      Temperature (℃)
    """
    while True:
        i = t.find(r'\g(')
        if i == -1:
            break
        j = t.find('))', i)
        if j == -1:
            break
        t = t[:i] + t[j + 2:]
    t = t.replace('°', '').replace('℃', 'K')
    t = t.replace(r'C\(41)', r'K\(41)')
    if 'K' not in t and 'C)' in t:
        t = t.replace('C)', 'K)')
    return t


def set_axis_kelvin(layer, axis='x', offset=273.15, title='Temperature (K)'):
    """Change an axis title to K and display tick labels as Celsius + offset.

    The data and axis limits are NOT changed; only the displayed labels move.
    """
    ax = layer.axis(axis)
    old = ax.title
    new = to_kelvin_title(old)
    if title:
        new = title
    ax.title = new
    layer.lt_exec(f'layer.{axis}.label.formula$ = "x+{offset:g}";')
    return old, new


# ---------------------------------------------------------------------------
# reading plot data
# ---------------------------------------------------------------------------

def read_plot_data(layer, plot_index):
    """Return [(x, y), ...] for a 1-based plot index in a layer.

    Uses the documented LabTalk route:
        range -wx rx = N;
        range -wy ry = N;
    Data inside a graph are references to worksheet/loose datasets.
    """
    layer.activate()
    op.lt_exec(f'range -wx rx = {plot_index};')
    op.lt_exec(f'range -wy ry = {plot_index};')
    n = op.lt_int('rx.GetSize()')
    xs = [op.lt_float(f'rx[{i}]') for i in range(1, n + 1)]
    ys = [op.lt_float(f'ry[{i}]') for i in range(1, n + 1)]
    return [(x, y) for x, y in zip(xs, ys)
            if not (math.isnan(x) or math.isnan(y))]


def read_all_plots(gp):
    """Collect every plot in every layer of a graph page."""
    out = []
    for li in range(len(gp)):
        layer = gp[li]
        layer.activate()
        for pi, pl in enumerate(layer.plot_list(), start=1):
            data = read_plot_data(layer, pi)
            out.append({
                'layer': li,
                'layer_name': layer.name,
                'plot': pi,
                'name': pl.name,
                'color': list(pl.color),
                'symbol': pl.symbol_kind,
                'n': len(data),
                'data': data,
            })
    return out


# ---------------------------------------------------------------------------
# fitting
# ---------------------------------------------------------------------------

def logistic(x, bottom, top, xc, k):
    return bottom + (top - bottom) / (1 + np.exp((x - xc) / k))


def exp_growth(x, a, b, c):
    return a + b * np.exp(c * x)


def exp_decay(x, a, b, c):
    return a + b * np.exp(-c * x)


def gaussian(x, y0, amp, xc, w):
    return y0 + amp * np.exp(-((x - xc) ** 2) / (2 * w * w))


def fit_series(x, y, model='logistic', nfit=551):
    """Fit x/y and return dict(popt, r2, xfit, yfit, model).

    Models: logistic, exp_growth, exp_decay, gaussian, poly2, poly3.
    """
    if np is None or curve_fit is None:
        raise RuntimeError('numpy/scipy not installed')
    x = np.asarray(x, dtype=float)
    y = np.asarray(y, dtype=float)
    if model == 'logistic':
        fn = logistic
        p0 = [min(y), max(y), float(np.median(x)), 50.0]
        bounds = ([0, 0, 0, 1], [10, 10, 700, 500])
    elif model == 'exp_growth':
        fn = exp_growth
        p0 = [float(min(y)), 1.0, 0.01]
        bounds = ([0, -1e3, 0], [1e4, 1e4, 1])
    elif model == 'exp_decay':
        fn = exp_decay
        p0 = [float(min(y)), float(max(y) - min(y)), 0.01]
        bounds = ([-1e3, 0, 0], [1e4, 1e4, 1])
    elif model == 'gaussian':
        fn = gaussian
        p0 = [float(min(y)), float(max(y) - min(y)), float(x[np.argmax(y)]), 80.0]
        bounds = ([-1e4, 0, 0, 1], [1e4, 1e4, 1e4, 1e4])
    elif model in ('poly2', 'poly3'):
        deg = int(model[-1])
        coeff = np.polyfit(x, y, deg)
        pred = np.polyval(coeff, x)
        r2 = 1 - ((y - pred) ** 2).sum() / ((y - y.mean()) ** 2).sum()
        xfit = np.linspace(float(x.min()), float(x.max()), nfit)
        return {'model': model, 'popt': coeff.tolist(), 'r2': float(r2),
                'xfit': xfit.tolist(), 'yfit': np.polyval(coeff, xfit).tolist()}
    else:
        raise ValueError(f'unknown model: {model}')

    popt, _ = curve_fit(fn, x, y, p0=p0, bounds=bounds, maxfev=100000)
    pred = fn(x, *popt)
    r2 = 1 - ((y - pred) ** 2).sum() / ((y - y.mean()) ** 2).sum()
    xfit = np.linspace(float(x.min()), float(x.max()), nfit)
    return {'model': model, 'popt': [float(v) for v in popt],
            'r2': float(r2), 'xfit': xfit.tolist(),
            'yfit': [float(v) for v in fn(xfit, *popt)]}


# ---------------------------------------------------------------------------
# adding fit curves / envelope / legend to an Origin graph
# ---------------------------------------------------------------------------

def make_fit_sheet(xfit, columns, name='FitData'):
    """Create a worksheet with X + one column per fit series.

    columns: dict[str, list[float]] (same length as xfit)
    Returns the worksheet.
    """
    try:
        old = op.find_sheet('w', name)
        if old:
            old.destroy()
    except Exception:
        pass
    wks = op.new_sheet('w', name)
    wks.from_list(0, list(xfit), lname='X')
    for i, (label, values) in enumerate(columns.items(), start=1):
        wks.from_list(i, list(values), lname=label)
    return wks


def add_fit_line(layer, wks, ycol, color=None, width=2):
    """Add a native Origin line plot from a worksheet column."""
    p = layer.add_plot(wks, ycol, 0, type='l')
    if color is not None:
        p.color = color
    try:
        p.set_cmd(f'-w {width}')
    except Exception:
        pass
    return p


def add_envelope(layer, wks, low_col, high_col, color_index=33547976):
    """Fill between two envelope curves.

    color_index 33547976 = light blue from color(200,230,255).
    The fill is applied to the lower curve and fills to the next plot.
    """
    low = layer.add_plot(wks, low_col, 0, type='l')
    high = layer.add_plot(wks, high_col, 0, type='l')
    for cmd in ['-pfv 8', '-pf 1', '-pfn 1', f'-cf {color_index}']:
        try:
            low.set_cmd(cmd)
        except Exception:
            pass
    for p in (low, high):
        try:
            p.set_cmd('-c 5')
            p.set_cmd('-w 1')
        except Exception:
            pass
    return low, high


def remove_auto_legend(layer):
    """Delete Origin's auto legend so added fit plots do not pollute it."""
    try:
        layer.remove_label('Legend')
    except Exception:
        pass


def add_static_legend(layer, entries, x=340, y=2.9, fontsize=20, dy=-0.17):
    """Add static text entries, e.g. [('● GSD Cu', (255,0,0)), ...]."""
    labels = []
    for i, (text, color) in enumerate(entries):
        lab = layer.add_label(text, x, y + i * dy)
        lab.color = color
        try:
            lab.set_float('fontsize', fontsize)
        except Exception:
            pass
        labels.append(lab)
    return labels


def save_export(gp, out_opju, png=None, overwrite=True):
    if overwrite and os.path.exists(out_opju):
        os.remove(out_opju)
    ok = bool(op.save(out_opju))
    if png:
        gp.activate()
        gp.save_fig(png)
    return ok


# ---------------------------------------------------------------------------
# CLI
# ---------------------------------------------------------------------------

def _cli():
    ap = argparse.ArgumentParser(description='Origin OPJU operations')
    sub = ap.add_subparsers(dest='cmd', required=True)

    p = sub.add_parser('read', help='read all plot data to JSON')
    p.add_argument('--opju', required=True)
    p.add_argument('--json', required=True)

    p = sub.add_parser('axis', help='convert a Celsius axis title/ticks to K')
    p.add_argument('--opju', required=True)
    p.add_argument('--axis', default='x')
    p.add_argument('--out', required=True)
    p.add_argument('--png')

    p = sub.add_parser('fit', help='fit one plot and add the curve')
    p.add_argument('--opju', required=True)
    p.add_argument('--plot', type=int, required=True)
    p.add_argument('--model', default='logistic',
                   choices=['logistic', 'exp_growth', 'exp_decay', 'gaussian', 'poly2', 'poly3'])
    p.add_argument('--out', required=True)
    p.add_argument('--png')
    p.add_argument('--color', default='255,0,0')

    a = ap.parse_args()
    try:
        if a.cmd == 'read':
            connect_standalone()
            gp = open_project(a.opju)
            data = {gp[0].name: read_all_plots(gp[0])}
            with open(a.json, 'w', encoding='utf-8') as f:
                json.dump(data, f, ensure_ascii=False, indent=2)
            print(json.dumps({k: len(v) for k, v in data.items()}, ensure_ascii=False))
        elif a.cmd == 'axis':
            connect_standalone()
            gp = open_project(a.opju)
            gl = gp[0]
            gl.activate()
            print(set_axis_kelvin(gl, a.axis))
            print('saved', save_export(gp, a.out, a.png))
        elif a.cmd == 'fit':
            connect_standalone()
            gp = open_project(a.opju)
            gl = gp[0]
            gl.activate()
            xy = read_plot_data(gl, a.plot)
            x = [p[0] for p in xy]
            y = [p[1] for p in xy]
            fit = fit_series(x, y, a.model)
            wks = make_fit_sheet(fit['xfit'], {a.model: fit['yfit']})
            color = tuple(int(v) for v in a.color.split(','))
            add_fit_line(gl, wks, 1, color=color, width=2)
            print('R2', fit['r2'], 'params', fit['popt'])
            print('saved', save_export(gp, a.out, a.png))
    finally:
        exit_origin()


if __name__ == '__main__':
    _cli()
