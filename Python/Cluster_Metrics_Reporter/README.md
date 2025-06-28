# 📊 Cluster Metrics Reporter

**Automated PowerPoint presentation generator for cluster performance metrics**

![PowerPoint](https://img.shields.io/badge/-PowerPoint-blue?logo=microsoft-powerpoint)
![Python](https://img.shields.io/badge/-Python-yellow?logo=python)
![Pandas](https://img.shields.io/badge/-Pandas-blue?logo=pandas)
![Logging](https://img.shields.io/badge/-Logging-lightgrey)

## 🚀 Overview

This Python script automatically generates a professional PowerPoint presentation with cluster performance metrics including:

- CPU usage (%)
- RAM consumption (GB)
- Network traffic (Mbps sent/received)

The tool processes Excel data and corresponding graphs, creating one slide per cluster with beautifully formatted metrics.

## ✨ Features

- **Automatic data processing** from Excel files
- **Smart image handling** with auto-scaling and positioning
- **Period comparison** (day/night/weekends)
- **Professional design** with:
  - Consistent styling
  - Corporate logo placement
  - Responsive layout
- **Comprehensive logging** for troubleshooting

## 📂 File Structure

project_root/
├── 📄 cluster_reporter.py # Main script
├── 📂 logs/ # Auto-generated log files
├── 📂 test-2/ # Sample data
│ ├── 📂 cpu/
│ ├── 📂 Оперативная память/
│ ├── 📂 Сеть (отправлено)/
│ └── 📂 Сеть (получено)/
├── 📂 logotip/ # Logo folder
│ └── 🖼️ logo.png
└── 📄 cluster_report.pptx # Output presentation

## 🛠️ Dependencies

Python 3.7+
Required packages:

## 🛠️ Установка

1. Клонируйте репозиторий:

```bash
git clone https://github.com/your/repo.git
```

2. Установите зависимости:

```bash
pip install -r requirements.txt
```

## 🚀 Пример быстрого старта

Основной скрипт:

```bash
# Для Linux/macOS
git clone https://github.com/yourusername/cluster-metrics-reporter.git && \
cd cluster-metrics-reporter && \
python3 -m venv venv && \
source venv/bin/activate && \
pip install -r requirements.txt && \
python cluster_reporter.py
```

````python
#!/usr/bin/env python3

# -_- coding: utf-8 -_-

"""
📊 Cluster Metrics Reporter
Automated PowerPoint presentation generator for cluster performance metrics

Features:

- Processes CPU, RAM, and Network metrics from Excel files
- Generates one slide per cluster with metrics and graphs
- Auto-scales images and handles layout dynamically
- Supports day/night/weekends period comparison
  """

import logging
from datetime import datetime
from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.dml.color import RGBColor
import pandas as pd
import os
from PIL import Image

# ==================== CONFIGURATION ====================

# Path settings

LOG_DIR = r"C:\Users\all-sp-111\Desktop\My_Project\Python\Project\Slider\logs"
LOGO_PATH = r"C:\Users\all-sp-111\Desktop\My_Project\Python\Project\Slider\logotip\logo.png"
OUTPUT_PPTX = r"C:\Users\all-sp-111\Desktop\My_Project\Python\Project\Slider\cluster_report.pptx"

# Logo display settings

LOGO_SCALE = 0.22 # Scale factor relative to slide width
LOGO_RIGHT_MARGIN = Inches(0.12)
LOGO_TOP_MARGIN = Inches(0.12)

# Metric display styles

METRIC_STYLES = {
"header": {
'size': Pt(14),
'bold': True,
'color': RGBColor(0, 0, 0) # Black
},
"values": {
'size': Pt(12),
'bold': False,
'color': RGBColor(0, 0, 0) # Black
}
}

# Slide title style

TITLE_STYLE = {
'size': Pt(28),
'bold': True,
'color': RGBColor(0, 32, 96) # Dark blue
}

# Metric configurations

project_root = os.path.dirname(os.path.abspath(**file**))

METRICS_CONFIG = {
"CPU": {
"excel": os.path.join(project_root, r"test-2\cpu\cpu.xlsx"),
"image_folder": os.path.join(project_root, r"test-2\cpu"),
"unit": "%",
"img_width": 10.5, # Inches
"textbox_height": Inches(1.2),
},
"RAM": {
"excel": os.path.join(project_root, r"test-2\Оперативная память\Оперативная память.xlsx"),
"image_folder": os.path.join(project_root, r"test-2\Оперативная память"),
"unit": "GB",
"img_width": 10.5,
"textbox_height": Inches(1.2),
},
"Network_sent": {
"excel": os.path.join(project_root, r"test-2\Сеть (отправлено)\Сеть (отправлено).xlsx"),
"image_folder": os.path.join(project_root, r"test-2\Сеть (отправлено)"),
"unit": "Mbps",
"img_width": 10.5,
"textbox_height": Inches(1.2),
"max_img_height": Inches(3.5)
},
"Network_received": {
"excel": os.path.join(project_root, r"test-2\Сеть (получено)\Сеть (получено).xlsx"),
"unit": "Mbps",
"textbox_height": Inches(1.2)
}
}

# ================= END CONFIGURATION ==================

def setup_logging():
"""Configure logging system with file and console output"""
if not os.path.exists(LOG_DIR):
os.makedirs(LOG_DIR)

    log_file = os.path.join(
        LOG_DIR,
        f"report_{datetime.now().strftime('%Y%m%d_%H%M%S')}.log"
    )

    logging.basicConfig(
        level=logging.INFO,
        format='%(asctime)s - %(levelname)s - %(message)s',
        handlers=[
            logging.FileHandler(log_file),
            logging.StreamHandler()
        ]
    )
    return log_file

def add_logo_to_slide(slide, prs):
"""Add logo to the top-right corner of slide"""
if not os.path.exists(LOGO_PATH):
logging.warning(f"Logo not found: {LOGO_PATH}")
return

    try:
        with Image.open(LOGO_PATH) as img:
            # Calculate logo dimensions maintaining aspect ratio
            img_width, img_height = img.size
            logo_width = int(prs.slide_width * LOGO_SCALE)
            logo_height = int(logo_width * img_height / img_width)

            # Position calculation
            left = prs.slide_width - logo_width - LOGO_RIGHT_MARGIN
            top = LOGO_TOP_MARGIN

            slide.shapes.add_picture(
                LOGO_PATH,
                left, top,
                width=logo_width,
                height=logo_height
            )
    except Exception as e:
        logging.error(f"Error adding logo: {str(e)}")

def get*clusters_from_images(image_folder):
"""Extract cluster names from image filenames"""
clusters = set()
if os.path.exists(image_folder):
for file in os.listdir(image_folder):
if file.endswith('\_grafic.png'): # Handle different filename patterns
if file.startswith(('gp*', 'gp-')):
cluster = file[3:-11]
else:
cluster = file[:-11]
clusters.add(cluster)
logging.info(f"Found cluster from image: {cluster}")
return sorted(clusters)

def load_metric_data(file_path, metric_name):
"""Load and process metric data from Excel file"""
try:
df = pd.read_excel(file_path)

        # Reshape from wide to long format
        value_vars = [col for col in df.columns if col != 'period']
        df = pd.melt(
            df,
            id_vars=['period'],
            value_vars=value_vars,
            var_name='cluster',
            value_name='value'
        )

        # Data cleaning
        df['period'] = df['period'].astype(str).str.strip().str.lower()
        df['value'] = pd.to_numeric(
            df['value'].astype(str).str.replace(',', '.'),
            errors='coerce'
        )
        df = df.dropna(subset=['value', 'cluster'])

        # Pivot to cluster-period matrix
        result = df.pivot(index='cluster', columns='period', values='value')

        logging.info(f"Successfully loaded {len(result)} clusters for {metric_name}")
        return result
    except Exception as e:
        logging.error(f"Error loading {file_path}: {str(e)}")
        return pd.DataFrame()

def get_cluster_values(df, cluster_name, metric_name):
"""Get metric values for specific cluster across periods"""
if df.empty:
return {'day': 'N/A', 'night': 'N/A', 'weekends': 'N/A'}

    # Find cluster name matches (case insensitive)
    cluster_matches = [
        c for c in df.index
        if str(cluster_name).lower() in str(c).lower()
    ]

    if not cluster_matches:
        logging.warning(f"No matches for cluster {cluster_name} in {metric_name}")
        return {'day': 'N/A', 'night': 'N/A', 'weekends': 'N/A'}

    matched_cluster = cluster_matches[0]
    values = {
        'day': f"{df.loc[matched_cluster, 'day']:.2f}".replace('.', ',') if 'day' in df.columns else 'N/A',
        'night': f"{df.loc[matched_cluster, 'night']:.2f}".replace('.', ',') if 'night' in df.columns else 'N/A',
        'weekends': f"{df.loc[matched_cluster, 'weekends']:.2f}".replace('.', ',') if 'weekends' in df.columns else 'N/A'
    }

    return values

def create_presentation(metric_data):
"""Generate PowerPoint presentation with cluster metrics""" # Initialize presentation
prs = Presentation()
prs.slide_width = Inches(13.33) # 16:9 aspect ratio
prs.slide_height = Inches(7.5)

    # Get all unique clusters from image files
    all_clusters = set()
    for metric, config in METRICS_CONFIG.items():
        if "image_folder" in config:
            clusters = get_clusters_from_images(config["image_folder"])
            all_clusters.update(clusters)

    if not all_clusters:
        logging.error("No clusters found in image files!")
        return

    # Create one slide per cluster
    for cluster in sorted(all_clusters):
        slide = prs.slides.add_slide(prs.slide_layouts[5])  # Blank layout
        add_logo_to_slide(slide, prs)

        # Set slide title
        title = slide.shapes.title
        title.text = f"Кластер: {cluster}"
        title.text_frame.paragraphs[0].font.bold = TITLE_STYLE['bold']
        title.text_frame.paragraphs[0].font.size = TITLE_STYLE['size']
        title.text_frame.paragraphs[0].font.color.rgb = TITLE_STYLE['color']

        # Position elements
        left = Inches(0.3)   # Left margin
        top = Inches(1.5)    # Start below title
        text_width = Inches(2.2)  # Width for metric text

        # Add metrics for each cluster
        for metric in [m for m in METRICS_CONFIG.keys() if m != "Network_received"]:
            config = METRICS_CONFIG[metric]
            df = metric_data.get(metric, pd.DataFrame())

            # For Network_sent, also show Network_received
            df2 = metric_data.get("Network_received", pd.DataFrame()) if metric == "Network_sent" else None

            values = get_cluster_values(df, cluster, metric)

            # Create text box for metrics
            textbox = slide.shapes.add_textbox(left, top, text_width, config["textbox_height"])
            tf = textbox.text_frame
            tf.word_wrap = True

            # Metric header
            p = tf.add_paragraph()
            p.text = metric.replace("_", " ")
            p.font.bold = METRIC_STYLES["header"]['bold']
            p.font.size = METRIC_STYLES["header"]['size']
            p.font.color.rgb = METRIC_STYLES["header"]['color']

            # Period values
            for period in ['day', 'night', 'weekends']:
                if values[period] != 'N/A':
                    p = tf.add_paragraph()
                    p.text = f"{period.capitalize()}: {values[period]} {config['unit']}"
                    p.font.size = METRIC_STYLES["values"]['size']
                    p.font.color.rgb = METRIC_STYLES["values"]['color']

            # Add Network_received to same textbox if showing Network_sent
            if metric == "Network_sent" and df2 is not None:
                values2 = get_cluster_values(df2, cluster, "Network_received")
                if any(v != 'N/A' for v in values2.values()):
                    p = tf.add_paragraph()
                    p.text = "Network received:"
                    p.font.bold = METRIC_STYLES["header"]['bold']
                    p.font.size = METRIC_STYLES["header"]['size']
                    p.font.color.rgb = METRIC_STYLES["header"]['color']

                    for period in ['day', 'night', 'weekends']:
                        if values2[period] != 'N/A':
                            p = tf.add_paragraph()
                            p.text = f"{period.capitalize()}: {values2[period]} {METRICS_CONFIG['Network_received']['unit']}"
                            p.font.size = METRIC_STYLES["values"]['size']
                            p.font.color.rgb = METRIC_STYLES["values"]['color']

            # Add corresponding graph image
            if "image_folder" in config:
                # Try different filename patterns
                img_patterns = [
                    f"{cluster}_grafic.png",
                    f"gp_{cluster}_grafic.png",
                    f"gp-{cluster}_grafic.png"
                ]

                img_path = next(
                    (os.path.join(config["image_folder"], f)
                    for f in img_patterns
                    if os.path.exists(os.path.join(config["image_folder"], f))
                ) if os.path.exists(config["image_folder"]) else None

                if img_path:
                    try:
                        with Image.open(img_path) as img:
                            # Calculate image dimensions maintaining aspect ratio
                            img_width = Inches(config["img_width"])
                            aspect = img.width / img.height
                            img_height = img_width / aspect

                            # Apply height limit for Network_sent
                            if "max_img_height" in config:
                                img_height = min(img_height, config["max_img_height"])

                            img_left = left + text_width + Inches(0.1)

                            # Adjust if exceeds slide width
                            if img_left + img_width > prs.slide_width - Inches(0.3):
                                img_width = prs.slide_width - img_left - Inches(0.3)
                                img_height = img_width / aspect

                            # Adjust if exceeds slide height
                            available_height = prs.slide_height - top - Inches(0.5)
                            if img_height > available_height:
                                img_height = available_height
                                img_width = img_height * aspect

                            # Add image to slide
                            slide.shapes.add_picture(
                                img_path,
                                img_left, top,
                                width=img_width,
                                height=img_height
                            )
                            top += max(config["textbox_height"], img_height) + Inches(0.2)
                    except Exception as e:
                        logging.error(f"Image error {img_path}: {str(e)}")
                        slide.shapes.add_textbox(
                            left + text_width + Inches(0.1),
                            top,
                            img_width,
                            Inches(0.8)
                        ).text_frame.text = "Image Error"
                        top += Inches(1.0)
                else:
                    slide.shapes.add_textbox(
                        left + text_width + Inches(0.1),
                        top,
                        img_width,
                        Inches(0.8)
                    ).text_frame.text = "No Graph"
                    top += Inches(1.0)
            else:
                top += config["textbox_height"] + Inches(0.2)

    # Save final presentation
    try:
        prs.save(OUTPUT_PPTX)
        logging.info(f"Presentation saved: {OUTPUT_PPTX}")
    except Exception as e:
        logging.error(f"Save error: {str(e)}")

if **name** == "**main**":
log_file = setup_logging()
logging.info(f"Script started. Log file: {log_file}")

    # Validate input files
    missing_files = []
    for metric, config in METRICS_CONFIG.items():
        if not os.path.exists(config["excel"]):
            missing_files.append(f"{metric} file: {config['excel']}")
        if "image_folder" in config and not os.path.exists(config["image_folder"]):
            missing_files.append(f"{metric} folder: {config['image_folder']}")

    if missing_files:
        logging.error("Missing files/folders:")
        for item in missing_files:
            logging.error(f" - {item}")
    else:
        # Load all metric data
        metric_data = {}
        for metric, config in METRICS_CONFIG.items():
            logging.info(f"\nLoading {metric} data...")
            data = load_metric_data(config["excel"], metric)
            if not data.empty:
                metric_data[metric] = data
                logging.info(f"Loaded {len(data)} clusters for {metric}")
            else:
                logging.error(f"Failed to load {metric} data")

        if metric_data:
            logging.info("\nCreating presentation...")
            create_presentation(metric_data)
        else:
            logging.error("No data available for presentation")

    logging.info("Script completed")

````
