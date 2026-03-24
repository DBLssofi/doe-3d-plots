import pandas as pd
import numpy as np
import matplotlib.pyplot as plt
import seaborn as sns
from mpl_toolkits.mplot3d import Axes3D
from scipy.interpolate import Rbf
from pandas.plotting import parallel_coordinates
import plotly.express as px
import plotly.graph_objects as go
import os

# ============================================================
# LOAD DATA
# ============================================================
file_path = r"G:\My Drive\Chemical engineering\Project work\WP2\DOE\python code\data set for python code.xlsx"
df = pd.read_excel(file_path)
df.columns = ['Temperature', 'pH', 'Buffer', 'RemainingActivity', 'StartingActivity']

# ============================================================
# OUTPUT FOLDER FOR INTERACTIVE HTML PLOTS
# ============================================================
output_dir = "3d_plots"
os.makedirs(output_dir, exist_ok=True)
print(f"HTML plots will be saved to: {os.path.abspath(output_dir)}")

# ============================================================
# HELPER: RBF GRID INTERPOLATION
# Used by both matplotlib and Plotly surface plots
# ============================================================
def make_rbf_grid(subset, response, grid_size=100):
    """Returns Xi, Yi, Zi mesh for a given subset and response column."""
    x = subset['Temperature'].values
    y = subset['pH'].values
    z = subset[response].values
    xi = np.linspace(x.min(), x.max(), grid_size)
    yi = np.linspace(y.min(), y.max(), grid_size)
    Xi, Yi = np.meshgrid(xi, yi)
    rbf = Rbf(x, y, z, function='multiquadric', smooth=0.1)
    Zi = rbf(Xi, Yi)
    return x, y, z, Xi, Yi, Zi


# ============================================================
# PART 1A: MATPLOTLIB SMOOTH 3D SURFACES
# ============================================================
def plot_smooth_surface_mpl(response, buffer_level):
    subset = df[df['Buffer'] == buffer_level]
    if len(subset) < 4:
        print(f"Skipping: Not enough points for Buffer={buffer_level}")
        return

    x, y, z, Xi, Yi, Zi = make_rbf_grid(subset, response)

    fig = plt.figure(figsize=(12, 8))
    ax = fig.add_subplot(111, projection='3d')
    surf = ax.plot_surface(Xi, Yi, Zi, cmap='viridis', edgecolor='none', alpha=0.9, antialiased=True)
    ax.scatter(x, y, z, color='black', s=60, edgecolors='white', linewidth=1.5, label='Actual Data', zorder=5)
    ax.set_xlabel('Temperature')
    ax.set_ylabel('pH')
    ax.set_zlabel(response)
    ax.set_title(f'3D Surface: {response}  |  Buffer = {buffer_level}')
    fig.colorbar(surf, ax=ax, shrink=0.5, aspect=10)
    ax.legend()
    plt.tight_layout()
    plt.show()

print("\n--- Matplotlib 3D Surfaces ---")
for b in df['Buffer'].unique():
    plot_smooth_surface_mpl('RemainingActivity', b)
    plot_smooth_surface_mpl('StartingActivity', b)


# ============================================================
# PART 1B: MATPLOTLIB GLOBAL 3D SCATTER
# ============================================================
print("\n--- Matplotlib Global 3D Scatter ---")
fig = plt.figure(figsize=(10, 8))
ax = fig.add_subplot(111, projection='3d')
sc = ax.scatter(
    df['Temperature'], df['pH'], df['Buffer'],
    c=df['RemainingActivity'], cmap='plasma', s=80, edgecolors='w'
)
ax.set_xlabel('Temperature')
ax.set_ylabel('pH')
ax.set_zlabel('Buffer')
plt.title("3D Global Scatter — Remaining Activity")
plt.colorbar(sc, label='Remaining Activity')
plt.tight_layout()
plt.show()


# ============================================================
# PART 2: SMOOTH 2D CONTOUR PLOTS
# ============================================================
def plot_smooth_contour(response, buffer_level):
    subset = df[df['Buffer'] == buffer_level]
    if len(subset) < 4:
        return

    x, y, z, Xi, Yi, Zi = make_rbf_grid(subset, response, grid_size=200)

    plt.figure(figsize=(8, 6))
    cp = plt.contourf(Xi, Yi, Zi, levels=30, cmap='RdYlBu_r')
    plt.colorbar(cp, label=response)
    plt.scatter(x, y, color='black', edgecolors='white', s=60, zorder=5)
    plt.xlabel('Temperature')
    plt.ylabel('pH')
    plt.title(f'Smooth Contour: {response}  |  Buffer = {buffer_level}')
    plt.tight_layout()
    plt.show()

print("\n--- Contour Plots ---")
for b in df['Buffer'].unique():
    plot_smooth_contour('RemainingActivity', b)


# ============================================================
# PART 3: HEATMAPS
# ============================================================
print("\n--- Heatmaps ---")
for b in df['Buffer'].unique():
    subset = df[df['Buffer'] == b]
    try:
        pivot = subset.pivot_table(index='Temperature', columns='pH', values='RemainingActivity')
        plt.figure(figsize=(6, 4))
        sns.heatmap(pivot, annot=True, cmap='YlGnBu', fmt=".2f")
        plt.title(f'Heatmap: Remaining Activity  |  Buffer = {b}')
        plt.tight_layout()
        plt.show()
    except Exception as e:
        print(f"Heatmap skipped for Buffer={b}: {e}")


# ============================================================
# PART 4: PARALLEL COORDINATES
# ============================================================
print("\n--- Parallel Coordinates ---")
df_parallel = df.copy()
df_parallel['Buffer'] = df_parallel['Buffer'].astype(str)
plt.figure(figsize=(12, 5))
parallel_coordinates(
    df_parallel, class_column='Buffer',
    cols=['Temperature', 'pH', 'RemainingActivity', 'StartingActivity'],
    colormap='Set1'
)
plt.title("Parallel Coordinates Analysis")
plt.tight_layout()
plt.show()


# ============================================================
# PART 5: CORRELATION SCATTER
# ============================================================
print("\n--- Correlation Scatter ---")
plt.figure(figsize=(10, 6))
sns.scatterplot(
    data=df, x='StartingActivity', y='RemainingActivity',
    hue='Temperature', size='Buffer', palette='viridis', sizes=(50, 400)
)
plt.title("Starting vs Remaining Activity Correlation")
plt.legend(bbox_to_anchor=(1.05, 1), loc=2)
plt.tight_layout()
plt.show()


# ============================================================
# PART 6: PLOTLY INTERACTIVE 3D — EXPORT TO HTML
# These are the files you will host on GitHub Pages and
# embed directly into your Canva presentation.
# ============================================================

print("\n--- Exporting Interactive Plotly HTML Files ---")

# --- 6.1 Global 3D Scatter ---
fig_scatter = px.scatter_3d(
    df,
    x='Temperature', y='pH', z='Buffer',
    color='RemainingActivity',
    size='StartingActivity',
    color_continuous_scale='Plasma',
    title="Interactive 3D DOE — Remaining Activity"
)
fig_scatter.update_layout(
    scene=dict(
        xaxis_title='Temperature',
        yaxis_title='pH',
        zaxis_title='Buffer'
    ),
    margin=dict(l=0, r=0, t=40, b=0)
)
scatter_path = os.path.join(output_dir, "global_scatter.html")
fig_scatter.write_html(scatter_path)
print(f"  Saved: {scatter_path}")


# --- 6.2 Smooth Surface Plots per Buffer ---
def export_surface_plotly(response, buffer_level):
    subset = df[df['Buffer'] == buffer_level]
    if len(subset) < 4:
        print(f"  Skipping surface export: not enough points for Buffer={buffer_level}")
        return

    x, y, z, Xi, Yi, Zi = make_rbf_grid(subset, response)

    fig = go.Figure()

    # Smooth interpolated surface
    fig.add_trace(go.Surface(
        x=Xi, y=Yi, z=Zi,
        colorscale='Viridis',
        opacity=0.88,
        name='Surface',
        showscale=True,
        contours=dict(
            z=dict(show=True, usecolormap=True, highlightcolor="white", project_z=True)
        )
    ))

    # Actual DOE data points
    fig.add_trace(go.Scatter3d(
        x=x, y=y, z=z,
        mode='markers',
        marker=dict(size=6, color='black', line=dict(color='white', width=1)),
        name='DOE Data Points'
    ))

    fig.update_layout(
        title=f'{response} — Buffer = {buffer_level}',
        scene=dict(
            xaxis_title='Temperature',
            yaxis_title='pH',
            zaxis_title=response,
            camera=dict(eye=dict(x=1.5, y=1.5, z=1.2))
        ),
        margin=dict(l=0, r=0, t=50, b=0),
        legend=dict(x=0.01, y=0.99)
    )

    # Sanitize filename (remove decimals/spaces)
    buffer_str = str(buffer_level).replace('.', '_').replace(' ', '_')
    filename = f"surface_{response}_buffer{buffer_str}.html"
    filepath = os.path.join(output_dir, filename)
    fig.write_html(filepath)
    print(f"  Saved: {filepath}")

for b in df['Buffer'].unique():
    export_surface_plotly('RemainingActivity', b)
    export_surface_plotly('StartingActivity', b)


# ============================================================
# PART 7: GENERATE INDEX PAGE
# Upload this alongside the plots to GitHub Pages.
# It gives you a clickable menu of all your 3D plots.
# ============================================================

html_files = [f for f in os.listdir(output_dir) if f.endswith('.html') and f != 'index.html']
html_files.sort()

index_html = """<!DOCTYPE html>
<html lang="en">
<head>
  <meta charset="UTF-8" />
  <meta name="viewport" content="width=device-width, initial-scale=1.0"/>
  <title>DOE 3D Plots</title>
  <style>
    body {
      font-family: 'Segoe UI', sans-serif;
      background: #0f1117;
      color: #e0e0e0;
      max-width: 800px;
      margin: 60px auto;
      padding: 0 24px;
    }
    h1 { color: #7dd3fc; letter-spacing: 1px; }
    p  { color: #94a3b8; margin-bottom: 32px; }
    ul { list-style: none; padding: 0; }
    li { margin: 12px 0; }
    a {
      display: block;
      padding: 14px 20px;
      background: #1e293b;
      border: 1px solid #334155;
      border-radius: 8px;
      color: #7dd3fc;
      text-decoration: none;
      transition: background 0.2s;
    }
    a:hover { background: #273549; border-color: #7dd3fc; }
  </style>
</head>
<body>
  <h1>🔬 DOE 3D Interactive Plots</h1>
  <p>Click any plot below to open it. All plots are fully interactive — rotate, zoom, and hover to explore.</p>
  <ul>
"""

for f in html_files:
    label = f.replace('.html', '').replace('_', ' ').title()
    index_html += f'    <li><a href="{f}" target="_blank">📊 {label}</a></li>\n'

index_html += """  </ul>
</body>
</html>
"""

index_path = os.path.join(output_dir, "index.html")
with open(index_path, 'w', encoding='utf-8') as f:
    f.write(index_html)

print(f"\n  Saved index page: {index_path}")
print("\n✅ All done!")
print(f"\nNext steps:")
print(f"  1. Upload the entire '{output_dir}/' folder to your GitHub repository")
print(f"  2. Enable GitHub Pages (Settings → Pages → main branch)")
print(f"  3. Your index will be at:  https://YOUR_USERNAME.github.io/YOUR_REPO/")
print(f"  4. In Canva: Add → Embed → paste each plot URL")