#!/usr/bin/env python3
"""
AQUILA Sovereign Platform — USD Exporter
=========================================
Converts the Aquila Master BOM + engineering specs into a complete OpenUSD
stage (.usda) with procedurally generated geometry for every BOM component.

Output: aquila_sovereign_platform.usda
Compatible with NVIDIA Omniverse Kit, Isaac Sim, USDView, and Cesium.

Usage:
    python export_aquila_usd.py
    python export_aquila_usd.py --output platform.usda
    python export_aquila_usd.py --z-up   # Z-up for Isaac Sim
"""

from __future__ import annotations

import argparse
import math
from datetime import datetime, timezone
from pathlib import Path


# ---------------------------------------------------------------------------
# BOM Data — extracted from Aquila_Master_BOM_Reconciled_v2 7-5-26
# Each entry: id, name, subsystem, dims_mm (x,y,z or diam,height), weight_g, qty, layer
# ---------------------------------------------------------------------------

# Drone platform (Phase 1: Firefly)
DRONE_PLATFORM = {
    "id": "UAV-001",
    "name": "Firefly Drone Platform",
    "dims_mm": (1325, 1325, 260),   # length × width × height
    "weight_g": 13600,
    "qty": 1,
    "type": "hexacopter",           # 6-arm heavy-lift
    "arms": 6,
    "arm_length_mm": 580,
    "prop_diam_mm": 800,
}

# Phase 1 operational stack components (flying now)
PHASE1_COMPONENTS = [
    {"id": "PWR-001", "name": "Tattu Plus 22 LiPo 22Ah",     "subsystem": "Power",       "dims": (215, 90, 75),   "weight": 2300, "qty": 2, "pos": (-150, 0, -80),  "color": (0.1, 0.6, 0.1), "shape": "box"},
    {"id": "PWR-002", "name": "Bioenno 24V 20Ah LiFePO4",    "subsystem": "Power",       "dims": (180, 75, 160),  "weight": 2600, "qty": 1, "pos": (150, 0, -80),   "color": (0.1, 0.5, 0.2), "shape": "box"},
    {"id": "CMP-001", "name": "Jetson AGX Orin 64GB",        "subsystem": "Compute",     "dims": (110, 110, 72),  "weight": 1350, "qty": 1, "pos": (0, 60, 60),     "color": (0.2, 0.2, 0.8), "shape": "box"},
    {"id": "CMP-002", "name": "Gauntlet AGX Thor Carrier",   "subsystem": "Compute",     "dims": (155, 126, 5),   "weight": 300,  "qty": 1, "pos": (0, 60, 95),     "color": (0.15, 0.15, 0.6),"shape": "box"},
    {"id": "GPR-001", "name": "Zond Aero LF GPR 50MHz",      "subsystem": "Subsurface",  "dims": (260, 160, 40),  "weight": 1200, "qty": 1, "pos": (0, -120, -40),  "color": (0.3, 0.8, 0.3), "shape": "box"},
    {"id": "STB-001", "name": "Gremsy T3V3 3-Axis Gimbal",   "subsystem": "Structure",   "dims": (240, 200, 180), "weight": 1950, "qty": 1, "pos": (0, -80, -180),  "color": (0.5, 0.5, 0.5), "shape": "gimbal"},
    {"id": "STR-001", "name": "CF OmniSpectral Housing",     "subsystem": "Structure",   "dims": (300, 300, 270), "weight": 1800, "qty": 1, "pos": (0, 0, 0),       "color": (0.12, 0.12, 0.15),"shape": "cylinder"},
    {"id": "OPT-001", "name": "FLIR Hadron 640R (LWIR+RGB)", "subsystem": "EO/IR",       "dims": (44, 44, 32),    "weight": 45,   "qty": 1, "pos": (0, -100, -220), "color": (0.2, 0.2, 0.3), "shape": "box"},
    {"id": "OPT-002", "name": "Sony ILX-LR1 61MP",           "subsystem": "Photogrammetry","dims": (100, 75, 65), "weight": 3,    "qty": 1, "pos": (60, -100, -220),"color": (0.1, 0.1, 0.1), "shape": "box"},
    {"id": "OPT-003", "name": "Stereolabs ZED 2i",           "subsystem": "Stereo Vision","dims": (175, 30, 33),  "weight": 166,  "qty": 1, "pos": (-60, -100, -220),"color": (0.0, 0.6, 0.9), "shape": "box"},
    {"id": "NAV-001", "name": "Garmin LIDAR-Lite v4",        "subsystem": "NAV",         "dims": (20, 48, 40),    "weight": 80,   "qty": 1, "pos": (120, 100, 80),  "color": (0.0, 0.8, 0.3), "shape": "box"},
    {"id": "CHE-001", "name": "Hamamatsu QCL Laser 4.6um",   "subsystem": "Chemistry",   "dims": (120, 60, 40),   "weight": 450,  "qty": 1, "pos": (-80, -40, 40),  "color": (0.8, 0.2, 0.1), "shape": "box"},
    {"id": "CHE-002", "name": "Vigo PVI-4TE-5 MCT Detector", "subsystem": "Chemistry",   "dims": (40, 40, 30),    "weight": 250,  "qty": 1, "pos": (-80, -80, 40),  "color": (0.7, 0.3, 0.0), "shape": "box"},
    {"id": "CHE-003", "name": "B&K Type 4955 PAS Detector",  "subsystem": "Chemistry",   "dims": (50, 50, 70),    "weight": 220,  "qty": 1, "pos": (-80, 20, 40),   "color": (0.6, 0.4, 0.0), "shape": "box"},
    {"id": "CHE-004", "name": "Hamamatsu C12880MA uSpec",    "subsystem": "Chemistry",   "dims": (27, 13, 8),     "weight": 10,   "qty": 1, "pos": (-80, 60, 40),   "color": (0.5, 0.5, 0.0), "shape": "box"},
    {"id": "CHE-005", "name": "Seeed VOC Sensor SGP30",      "subsystem": "Chemistry",   "dims": (25, 20, 10),    "weight": 12,   "qty": 1, "pos": (80, 40, 40),    "color": (0.4, 0.6, 0.0), "shape": "box"},
    {"id": "AUD-001", "name": "ReSpeaker 4-Mic Array",       "subsystem": "Audio",       "dims": (70, 70, 12),    "weight": 80,   "qty": 1, "pos": (80, 100, 60),   "color": (0.2, 0.0, 0.4), "shape": "cylinder"},
    {"id": "AUD-002", "name": "SPH0645LM4H MEMS Mic",        "subsystem": "Audio",       "dims": (3.5, 2.6, 1),   "weight": 1,    "qty": 1, "pos": (80, 120, 60),   "color": (0.3, 0.0, 0.3), "shape": "box"},
    {"id": "ENV-001", "name": "Sensirion SGP30+SHT31",       "subsystem": "Env",         "dims": (20, 15, 5),     "weight": 6,    "qty": 1, "pos": (100, 40, 40),   "color": (0.0, 0.4, 0.6), "shape": "box"},
]

# Phase 2 — Firefly hybrid mount components
FIREFLY_MOUNT = [
    {"id": "PWR-FLY-001", "name": "48V Buck Converter",       "subsystem": "Power",   "dims": (120, 80, 50),  "weight": 450,  "qty": 1, "pos": (0, 0, 120),    "color": (0.6, 0.5, 0.0), "shape": "box"},
    {"id": "THR-FLY-001", "name": "CF Heat Shield",           "subsystem": "Thermal", "dims": (400, 350, 25), "weight": 280,  "qty": 1, "pos": (0, 0, 160),    "color": (0.4, 0.4, 0.5), "shape": "box"},
    {"id": "MNT-FLY-001", "name": "CNC 15mm Rail Clamps",     "subsystem": "Mount",   "dims": (80, 60, 40),   "weight": 320,  "qty": 4, "pos": (200, 200, 100),"color": (0.5, 0.5, 0.5), "shape": "box"},
    {"id": "STR-POD-001", "name": "Vibration Dampener Pod",   "subsystem": "Pod",     "dims": (200, 150, 80), "weight": 600,  "qty": 1, "pos": (0, 0, 80),     "color": (0.3, 0.3, 0.3), "shape": "box"},
    {"id": "STR-POD-002", "name": "goBILDA 15mm Grid Frame",  "subsystem": "Pod",     "dims": (600, 500, 400),"weight": 3200, "qty": 1, "pos": (0, 0, 0),      "color": (0.25, 0.25, 0.3),"shape": "frame"},
    {"id": "STR-POD-003", "name": "QCL Drop Winch Assy",      "subsystem": "Pod",     "dims": (250, 150, 120),"weight": 1200, "qty": 1, "pos": (0, -200, -100),"color": (0.4, 0.4, 0.2), "shape": "box"},
    {"id": "ACT-REV-001", "name": "REV UltraPlanetary Gearbox","subsystem": "Actuation","dims": (80, 60, 50),  "weight": 250,  "qty": 1, "pos": (150, -200, -100),"color": (0.8, 0.1, 0.1),"shape": "cylinder"},
    {"id": "ACT-REV-002", "name": "REV NEO 550 Motor",        "subsystem": "Actuation","dims": (35, 35, 35),   "weight": 142,  "qty": 1, "pos": (150, -200, -60),"color": (0.7, 0.1, 0.1),"shape": "cylinder"},
    {"id": "CTL-REV-001", "name": "REV SPARK MAX Controller", "subsystem": "Control", "dims": (70, 35, 15),   "weight": 80,   "qty": 1, "pos": (150, -150, -60),"color": (0.1, 0.1, 0.1), "shape": "box"},
]

# Bio-Intel DNA stack
BIO_STACK = [
    {"id": "BIO-001", "name": "Oxford MinION Mk1C",           "subsystem": "Bio",     "dims": (105, 23, 33),  "weight": 87,   "qty": 1, "pos": (-150, -60, 40), "color": (0.0, 0.7, 0.3), "shape": "box"},
    {"id": "BIO-002", "name": "Micro-Peristaltic Pump",       "subsystem": "Bio",     "dims": (60, 40, 30),   "weight": 120,  "qty": 1, "pos": (-150, -30, 40), "color": (0.0, 0.5, 0.2), "shape": "box"},
    {"id": "BIO-003", "name": "STaPLE Soil Extractor",        "subsystem": "Bio",     "dims": (200, 80, 80),  "weight": 400,  "qty": 1, "pos": (-200, 0, -60),  "color": (0.0, 0.4, 0.1), "shape": "box"},
    {"id": "BIO-004", "name": "PlanetVac Mini-Nozzle",        "subsystem": "Bio",     "dims": (150, 60, 60),  "weight": 280,  "qty": 1, "pos": (-200, 0, -120), "color": (0.1, 0.5, 0.0), "shape": "cylinder"},
]


# ---------------------------------------------------------------------------
# USD Generation helpers
# ---------------------------------------------------------------------------

def fmt_v3(arr):
    return f"({arr[0]:.2f}, {arr[1]:.2f}, {arr[2]:.2f})"

def fmt_v3i(arr):
    return f"({arr[0]}, {arr[1]}, {arr[2]})"

def fmt_color(c):
    return f"({c[0]:.3f}, {c[1]:.3f}, {c[2]:.3f})"


def box_mesh(name, dims, pos, color, bom_id="", subsystem="", weight=0, extra_attrs=None):
    """Generate a USD Mesh box prim."""
    w, h, d = dims
    x, y, z = pos
    lines = []
    lines.append(f'        def Mesh "{name}"')
    lines.append("        {")
    lines.append(f'            float3[] extent = [({-w/2:.1f}, {-h/2:.1f}, {-d/2:.1f}), ({w/2:.1f}, {h/2:.1f}, {d/2:.1f})]')
    lines.append('            int[] faceVertexCounts = [4, 4, 4, 4, 4, 4]')
    lines.append('            int[] faceVertexIndices = [0,1,2,3, 4,5,6,7, 0,4,5,1, 1,5,6,2, 2,6,7,3, 3,7,4,0]')
    lines.append(f'            point3f[] points = [({-w/2:.1f},{-h/2:.1f},{-d/2:.1f}), ({w/2:.1f},{-h/2:.1f},{-d/2:.1f}), ({w/2:.1f},{h/2:.1f},{-d/2:.1f}), ({-w/2:.1f},{h/2:.1f},{-d/2:.1f}), ({-w/2:.1f},{-h/2:.1f},{d/2:.1f}), ({w/2:.1f},{-h/2:.1f},{d/2:.1f}), ({w/2:.1f},{h/2:.1f},{d/2:.1f}), ({-w/2:.1f},{h/2:.1f},{d/2:.1f})]')
    lines.append(f'            matrix4d xformOp:transform = ( (1,0,0,0), (0,1,0,0), (0,0,1,0), ({x:.2f},{y:.2f},{z:.2f},1) )')
    lines.append('            token[] xformOpOrder = ["xformOp:transform"]')
    if bom_id:
        lines.append(f'            string bomId = "{bom_id}"')
    if subsystem:
        lines.append(f'            string subsystem = "{subsystem}"')
    if weight:
        lines.append(f'            double weightGrams = {weight}')
    if extra_attrs:
        for k, v in extra_attrs.items():
            lines.append(f'            {k} = {v}')
    lines.append(f'            rel material:binding = </AquilaSovereign/Materials/{name}_Mat>')
    lines.append("        }")
    # Material
    lines.append(f'        def Material "{name}_Mat"')
    lines.append("        {")
    lines.append(f'            token outputs:surface.connect = </AquilaSovereign/Materials/{name}_Mat/Shader.outputs:surface>')
    lines.append('            def Shader "Shader"')
    lines.append("            {")
    lines.append('                uniform token info:id = "UsdPreviewSurface"')
    lines.append(f'                color3f inputs:diffuseColor = {fmt_color(color)}')
    lines.append('                float inputs:roughness = 0.5')
    lines.append('                float inputs:metallic = 0.3')
    lines.append('                token outputs:surface')
    lines.append("            }")
    lines.append("        }")
    return lines


def cylinder_mesh(name, diam, height, pos, color, bom_id="", subsystem="", weight=0):
    """Generate a USD Mesh cylinder prim (12-sided)."""
    r = diam / 2
    x, y, z = pos
    segments = 12
    pts_bot = []
    pts_top = []
    for i in range(segments):
        ang = 2 * math.pi * i / segments
        px = r * math.cos(ang)
        py = r * math.sin(ang)
        pts_bot.append(f"({px:.2f},{py:.2f},{-height/2:.2f})")
        pts_top.append(f"({px:.2f},{py:.2f},{height/2:.2f})")
    # Side faces
    fvc = []
    fvi = []
    for i in range(segments):
        ni = (i + 1) % segments
        fvc.append(4)
        fvi.extend([i, ni, ni + segments, i + segments])
    # Bottom cap
    fvc.append(segments)
    fvi.extend(range(segments))
    # Top cap
    fvc.append(segments)
    fvi.extend(range(segments, segments * 2))

    lines = []
    lines.append(f'        def Mesh "{name}"')
    lines.append("        {")
    lines.append(f'            float3[] extent = [({-r:.1f}, {-r:.1f}, {-height/2:.1f}), ({r:.1f}, {r:.1f}, {height/2:.1f})]')
    lines.append(f'            int[] faceVertexCounts = {fvc}')
    lines.append(f'            int[] faceVertexIndices = {fvi}')
    lines.append(f'            point3f[] points = [{", ".join(pts_bot)}, {", ".join(pts_top)}]')
    lines.append(f'            matrix4d xformOp:transform = ( (1,0,0,0), (0,1,0,0), (0,0,1,0), ({x:.2f},{y:.2f},{z:.2f},1) )')
    lines.append('            token[] xformOpOrder = ["xformOp:transform"]')
    if bom_id:
        lines.append(f'            string bomId = "{bom_id}"')
    if subsystem:
        lines.append(f'            string subsystem = "{subsystem}"')
    if weight:
        lines.append(f'            double weightGrams = {weight}')
    lines.append(f'            rel material:binding = </AquilaSovereign/Materials/{name}_Mat>')
    lines.append("        }")
    lines.append(f'        def Material "{name}_Mat"')
    lines.append("        {")
    lines.append(f'            token outputs:surface.connect = </AquilaSovereign/Materials/{name}_Mat/Shader.outputs:surface>')
    lines.append('            def Shader "Shader"')
    lines.append("            {")
    lines.append('                uniform token info:id = "UsdPreviewSurface"')
    lines.append(f'                color3f inputs:diffuseColor = {fmt_color(color)}')
    lines.append('                float inputs:roughness = 0.6')
    lines.append('                token outputs:surface')
    lines.append("            }")
    lines.append("        }")
    return lines


def gimbal_mesh(name, dims, pos, color, bom_id="", subsystem="", weight=0):
    """Generate a 3-axis gimbal: 3 nested rings."""
    w, h, d = dims
    x, y, z = pos
    lines = []
    # Outer ring (yaw)
    r_outer = max(w, h) / 2
    r_inner = r_outer * 0.85
    tube_r = d / 4
    lines.append(f'        def Xform "{name}"')
    lines.append("        {")
    lines.append(f'            matrix4d xformOp:transform = ( (1,0,0,0), (0,1,0,0), (0,0,1,0), ({x:.2f},{y:.2f},{z:.2f},1) )')
    lines.append('            token[] xformOpOrder = ["xformOp:transform"]')
    lines.append(f'            string bomId = "{bom_id}"')
    lines.append(f'            string subsystem = "{subsystem}"')
    lines.append(f'            double weightGrams = {weight}')
    # Yaw ring (torus approximation as cylinder ring)
    for axis, (label, (rx, ry, rz, rot)) in enumerate([
        ("Yaw",   (r_outer, r_outer, tube_r, (0, 0, 0))),
        ("Pitch", (r_inner, r_inner, tube_r * 0.8, (0, 90, 0))),
        ("Roll",  (r_inner * 0.7, r_inner * 0.7, tube_r * 0.6, (90, 0, 0))),
    ]):
        lines.append(f'            def Mesh "{label}Ring"')
        lines.append("            {")
        lines.append(f'                float3[] extent = [({-rx:.1f}, {-ry:.1f}, {-rz:.1f}), ({rx:.1f}, {ry:.1f}, {rz:.1f})]')
        # Simple torus as a flattened cylinder ring
        segs = 16
        pts = []
        for i in range(segs):
            ang = 2 * math.pi * i / segs
            px = rx * math.cos(ang)
            py = ry * math.sin(ang)
            pts.append(f"({px:.2f},{py:.2f},0)")
            pts.append(f"({px*0.9:.2f},{py*0.9:.2f},0)")
        fvc = []
        fvi = []
        for i in range(segs):
            ni = (i + 1) % segs
            fvc.append(4)
            fvi.extend([i*2, ni*2, ni*2+1, i*2+1])
        lines.append(f'                int[] faceVertexCounts = {fvc}')
        lines.append(f'                int[] faceVertexIndices = {fvi}')
        lines.append(f'                point3f[] points = [{", ".join(pts)}]')
        lines.append(f'                double3 xformOp:rotateXYZ = {fmt_v3i(rot)}')
        lines.append('                token[] xformOpOrder = ["xformOp:rotateXYZ"]')
        lines.append(f'                rel material:binding = </AquilaSovereign/Materials/{name}_Mat>')
        lines.append("            }")
    lines.append("        }")
    # Material
    lines.append(f'        def Material "{name}_Mat"')
    lines.append("        {")
    lines.append(f'            token outputs:surface.connect = </AquilaSovereign/Materials/{name}_Mat/Shader.outputs:surface>')
    lines.append('            def Shader "Shader"')
    lines.append("            {")
    lines.append('                uniform token info:id = "UsdPreviewSurface"')
    lines.append(f'                color3f inputs:diffuseColor = {fmt_color(color)}')
    lines.append('                float inputs:metallic = 0.7')
    lines.append('                float inputs:roughness = 0.3')
    lines.append('                token outputs:surface')
    lines.append("            }")
    lines.append("        }")
    return lines


def frame_mesh(name, dims, pos, color, bom_id="", subsystem="", weight=0):
    """Generate a 15mm grid frame as a wireframe box structure."""
    w, h, d = dims
    x, y, z = pos
    lines = []
    lines.append(f'        def Xform "{name}"')
    lines.append("        {")
    lines.append(f'            matrix4d xformOp:transform = ( (1,0,0,0), (0,1,0,0), (0,0,1,0), ({x:.2f},{y:.2f},{z:.2f},1) )')
    lines.append('            token[] xformOpOrder = ["xformOp:transform"]')
    lines.append(f'            string bomId = "{bom_id}"')
    lines.append(f'            string subsystem = "{subsystem}"')
    lines.append(f'            double weightGrams = {weight}')
    # 12 edge bars
    bar_r = 7.5  # 15mm diameter / 2
    corners = [
        (-w/2, -h/2, -d/2), (w/2, -h/2, -d/2), (w/2, h/2, -d/2), (-w/2, h/2, -d/2),
        (-w/2, -h/2, d/2),  (w/2, -h/2, d/2),  (w/2, h/2, d/2),  (-w/2, h/2, d/2),
    ]
    edges = [
        (0,1), (1,2), (2,3), (3,0),  # bottom
        (4,5), (5,6), (6,7), (7,4),  # top
        (0,4), (1,5), (2,6), (3,7),  # verticals
    ]
    for ei, (a, b) in enumerate(edges):
        ca = corners[a]
        cb = corners[b]
        mid = ((ca[0]+cb[0])/2, (ca[1]+cb[1])/2, (ca[2]+cb[2])/2)
        length = math.sqrt(sum((cb[i]-ca[i])**2 for i in range(3)))
        lines.append(f'            def Mesh "Bar_{ei:02d}"')
        lines.append("            {")
        lines.append(f'                float3[] extent = [({-bar_r}, {-bar_r}, {-length/2:.1f}), ({bar_r}, {bar_r}, {length/2:.1f})]')
        lines.append('                int[] faceVertexCounts = [4, 4, 4, 4, 4, 4]')
        lines.append('                int[] faceVertexIndices = [0,1,2,3, 4,5,6,7, 0,4,5,1, 1,5,6,2, 2,6,7,3, 3,7,4,0]')
        lines.append(f'                point3f[] points = [({-bar_r},{-bar_r},{-length/2:.1f}),({bar_r},{-bar_r},{-length/2:.1f}),({bar_r},{bar_r},{-length/2:.1f}),({-bar_r},{bar_r},{-length/2:.1f}),({-bar_r},{-bar_r},{length/2:.1f}),({bar_r},{-bar_r},{length/2:.1f}),({bar_r},{bar_r},{length/2:.1f}),({-bar_r},{bar_r},{length/2:.1f})]')
        lines.append(f'                matrix4d xformOp:transform = ( (1,0,0,0), (0,1,0,0), (0,0,1,0), ({mid[0]:.2f},{mid[1]:.2f},{mid[2]:.2f},1) )')
        lines.append('                token[] xformOpOrder = ["xformOp:transform"]')
        lines.append(f'                rel material:binding = </AquilaSovereign/Materials/{name}_Mat>')
        lines.append("            }")
    lines.append("        }")
    lines.append(f'        def Material "{name}_Mat"')
    lines.append("        {")
    lines.append(f'            token outputs:surface.connect = </AquilaSovereign/Materials/{name}_Mat/Shader.outputs:surface>')
    lines.append('            def Shader "Shader"')
    lines.append("            {")
    lines.append('                uniform token info:id = "UsdPreviewSurface"')
    lines.append(f'                color3f inputs:diffuseColor = {fmt_color(color)}')
    lines.append('                float inputs:metallic = 0.6')
    lines.append('                float inputs:roughness = 0.4')
    lines.append('                token outputs:surface')
    lines.append("            }")
    lines.append("        }")
    return lines


def drone_platform_mesh(platform):
    """Generate the hexacopter drone platform: central body + 6 arms + 6 rotors."""
    lines = []
    w, h, d = platform["dims_mm"]
    arm_len = platform["arm_length_mm"]
    prop_d = platform["prop_diam_mm"]
    num_arms = platform["arms"]

    lines.append('    def Xform "DronePlatform"')
    lines.append("    {")
    lines.append(f'        string bomId = "{platform["id"]}"')
    lines.append(f'        string componentName = "{platform["name"]}"')
    lines.append(f'        double weightGrams = {platform["weight_g"]}')
    lines.append(f'        int armCount = {num_arms}')
    lines.append('')

    # Central hub — carbon fiber body
    hub_w = 200
    lines.append('        def Mesh "CentralHub"')
    lines.append("        {")
    lines.append(f'            float3[] extent = [({-hub_w/2}, {-hub_w/2}, {-d/2}), ({hub_w/2}, {hub_w/2}, {d/2})]')
    lines.append('            int[] faceVertexCounts = [4, 4, 4, 4, 4, 4]')
    lines.append('            int[] faceVertexIndices = [0,1,2,3, 4,5,6,7, 0,4,5,1, 1,5,6,2, 2,6,7,3, 3,7,4,0]')
    lines.append(f'            point3f[] points = [({-hub_w/2},{-hub_w/2},{-d/2}),({hub_w/2},{-hub_w/2},{-d/2}),({hub_w/2},{hub_w/2},{-d/2}),({-hub_w/2},{hub_w/2},{-d/2}),({-hub_w/2},{-hub_w/2},{d/2}),({hub_w/2},{-hub_w/2},{d/2}),({hub_w/2},{hub_w/2},{d/2}),({-hub_w/2},{hub_w/2},{d/2})]')
    lines.append('            rel material:binding = </AquilaSovereign/Materials/DroneBody_Mat>')
    lines.append("        }")
    lines.append('')

    # Arms + motors + propellers
    for i in range(num_arms):
        ang = 2 * math.pi * i / num_arms
        arm_end_x = math.cos(ang) * arm_len
        arm_end_y = math.sin(ang) * arm_len
        arm_end_z = 0
        mid_x = arm_end_x / 2
        mid_y = arm_end_y / 2

        # Arm tube
        arm_w = 40
        arm_h = 40
        arm_len_actual = math.sqrt(arm_end_x**2 + arm_end_y**2)
        arm_rot_z = math.degrees(ang)
        lines.append(f'        def Mesh "Arm_{i:02d}"')
        lines.append("        {")
        lines.append(f'            float3[] extent = [({-arm_len_actual/2:.1f}, {-arm_w/2}, {-arm_h/2}), ({arm_len_actual/2:.1f}, {arm_w/2}, {arm_h/2})]')
        lines.append('            int[] faceVertexCounts = [4, 4, 4, 4, 4, 4]')
        lines.append('            int[] faceVertexIndices = [0,1,2,3, 4,5,6,7, 0,4,5,1, 1,5,6,2, 2,6,7,3, 3,7,4,0]')
        lines.append(f'            point3f[] points = [({-arm_len_actual/2:.1f},{-arm_w/2},{-arm_h/2}),({arm_len_actual/2:.1f},{-arm_w/2},{-arm_h/2}),({arm_len_actual/2:.1f},{arm_w/2},{-arm_h/2}),({-arm_len_actual/2:.1f},{arm_w/2},{-arm_h/2}),({-arm_len_actual/2:.1f},{-arm_w/2},{arm_h/2}),({arm_len_actual/2:.1f},{-arm_w/2},{arm_h/2}),({arm_len_actual/2:.1f},{arm_w/2},{arm_h/2}),({-arm_len_actual/2:.1f},{arm_w/2},{arm_h/2})]')
        lines.append(f'            double3 xformOp:rotateXYZ = (0, 0, {arm_rot_z:.2f})')
        lines.append(f'            matrix4d xformOp:transform = ( (1,0,0,0), (0,1,0,0), (0,0,1,0), ({mid_x:.2f},{mid_y:.2f},{arm_end_z:.2f},1) )')
        lines.append('            token[] xformOpOrder = ["xformOp:transform", "xformOp:rotateXYZ"]')
        lines.append('            rel material:binding = </AquilaSovereign/Materials/DroneBody_Mat>')
        lines.append("        }")

        # Motor (cylinder at arm tip)
        lines.append(f'        def Mesh "Motor_{i:02d}"')
        lines.append("        {")
        motor_r = 30
        motor_h = 60
        segs = 8
        pts = []
        for j in range(segs):
            a = 2 * math.pi * j / segs
            pts.append(f"({motor_r*math.cos(a):.1f},{motor_r*math.sin(a):.1f},{-motor_h/2:.1f})")
            pts.append(f"({motor_r*math.cos(a):.1f},{motor_r*math.sin(a):.1f},{motor_h/2:.1f})")
        fvc = []
        fvi = []
        for j in range(segs):
            nj = (j + 1) % segs
            fvc.append(4)
            fvi.extend([j*2, nj*2, nj*2+1, j*2+1])
        fvc.append(segs)
        fvi.extend(range(0, segs*2, 2))
        fvc.append(segs)
        fvi.extend(range(1, segs*2, 2))
        lines.append(f'            float3[] extent = [({-motor_r}, {-motor_r}, {-motor_h/2}), ({motor_r}, {motor_r}, {motor_h/2})]')
        lines.append(f'            int[] faceVertexCounts = {fvc}')
        lines.append(f'            int[] faceVertexIndices = {fvi}')
        lines.append(f'            point3f[] points = [{", ".join(pts)}]')
        lines.append(f'            matrix4d xformOp:transform = ( (1,0,0,0), (0,1,0,0), (0,0,1,0), ({arm_end_x:.2f},{arm_end_y:.2f},{motor_h/2+d/2:.2f},1) )')
        lines.append('            token[] xformOpOrder = ["xformOp:transform"]')
        lines.append('            rel material:binding = </AquilaSovereign/Materials/Motor_Mat>')
        lines.append("        }")

        # Propeller (thin cylinder)
        lines.append(f'        def Mesh "Propeller_{i:02d}"')
        lines.append("        {")
        prop_r = prop_d / 2
        prop_h = 3
        pts_p = []
        for j in range(segs):
            a = 2 * math.pi * j / segs
            pts_p.append(f"({prop_r*math.cos(a):.1f},{prop_r*math.sin(a):.1f},{-prop_h/2})")
            pts_p.append(f"({prop_r*math.cos(a):.1f},{prop_r*math.sin(a):.1f},{prop_h/2})")
        fvc_p = []
        fvi_p = []
        for j in range(segs):
            nj = (j + 1) % segs
            fvc_p.append(4)
            fvi_p.extend([j*2, nj*2, nj*2+1, j*2+1])
        fvc_p.append(segs)
        fvi_p.extend(range(0, segs*2, 2))
        fvc_p.append(segs)
        fvi_p.extend(range(1, segs*2, 2))
        lines.append(f'            float3[] extent = [({-prop_r:.1f}, {-prop_r:.1f}, {-prop_h/2}), ({prop_r:.1f}, {prop_r:.1f}, {prop_h/2})]')
        lines.append(f'            int[] faceVertexCounts = {fvc_p}')
        lines.append(f'            int[] faceVertexIndices = {fvi_p}')
        lines.append(f'            point3f[] points = [{", ".join(pts_p)}]')
        lines.append(f'            matrix4d xformOp:transform = ( (1,0,0,0), (0,1,0,0), (0,0,1,0), ({arm_end_x:.2f},{arm_end_y:.2f},{motor_h+d/2+5:.2f},1) )')
        lines.append('            token[] xformOpOrder = ["xformOp:transform"]')
        lines.append('            rel material:binding = </AquilaSovereign/Materials/Prop_Mat>')
        # Animated rotation
        lines.append('            float xformOp:rotateZ = 0')
        lines.append('            token[] xformOpOrder = ["xformOp:transform", "xformOp:rotateZ"]')
        lines.append('            float xformOp:rotateZ.timeSamples = {')
        for frame in range(1, 901, 10):
            spin = (frame / 30.0 * 720) % 360  # 720 deg/sec = 2 rev/sec
            lines.append(f'                {frame}: {spin:.1f},')
        lines.append('            }')
        lines.append("        }")
        lines.append('')

    # Landing gear (4 legs)
    for i in range(4):
        ang = math.pi / 4 + i * math.pi / 2
        lx = math.cos(ang) * 80
        ly = math.sin(ang) * 80
        lines.append(f'        def Mesh "LandingGear_{i:02d}"')
        lines.append("        {")
        leg_r = 10
        leg_h = 80
        segs = 6
        pts = []
        for j in range(segs):
            a = 2 * math.pi * j / segs
            pts.append(f"({leg_r*math.cos(a):.1f},{leg_r*math.sin(a):.1f},{-leg_h/2:.1f})")
            pts.append(f"({leg_r*math.cos(a):.1f},{leg_r*math.sin(a):.1f},{leg_h/2:.1f})")
        fvc = []
        fvi = []
        for j in range(segs):
            nj = (j + 1) % segs
            fvc.append(4)
            fvi.extend([j*2, nj*2, nj*2+1, j*2+1])
        fvc.append(segs)
        fvi.extend(range(0, segs*2, 2))
        fvc.append(segs)
        fvi.extend(range(1, segs*2, 2))
        lines.append(f'            float3[] extent = [({-leg_r}, {-leg_r}, {-leg_h/2}), ({leg_r}, {leg_r}, {leg_h/2})]')
        lines.append(f'            int[] faceVertexCounts = {fvc}')
        lines.append(f'            int[] faceVertexIndices = {fvi}')
        lines.append(f'            point3f[] points = [{", ".join(pts)}]')
        lines.append(f'            matrix4d xformOp:transform = ( (1,0,0,0), (0,1,0,0), (0,0,1,0), ({lx:.2f},{ly:.2f},{-d/2-leg_h/2:.2f},1) )')
        lines.append('            token[] xformOpOrder = ["xformOp:transform"]')
        lines.append('            rel material:binding = </AquilaSovereign/Materials/DroneBody_Mat>')
        lines.append("        }")

    lines.append("    }")
    lines.append("")
    return lines


# ---------------------------------------------------------------------------
# Main builder
# ---------------------------------------------------------------------------

def build_usda(up_axis="Y"):
    ts = datetime.now(timezone.utc).strftime("%Y-%m-%dT%H:%M:%SZ")
    lines = []
    lines.append("#usda 1.0")
    lines.append("(")
    lines.append('    defaultPrim = "AquilaSovereign"')
    lines.append(f'    doc = "AQUILA Sovereign Platform — full BOM-to-USD conversion. Generated {ts}"')
    lines.append('    metersPerUnit = 0.001')  # mm → m
    lines.append(f'    upAxis = "{up_axis}"')
    lines.append('    framesPerSecond = 30')
    lines.append('    startTimeCode = 1')
    lines.append('    endTimeCode = 900')
    lines.append(")")
    lines.append("")

    # Root
    lines.append('def Xform "AquilaSovereign"')
    lines.append("{")
    lines.append('    string mission = "AQUILA Sovereign — Orphan-Well Leak Detection + Bio-Intel"')
    lines.append('    string bomVersion = "Reconciled v2 7-5-26"')
    lines.append('    string platform = "Firefly Heavy-Lift Hexacopter"')
    lines.append('    int totalComponents = 50')
    lines.append('    double totalWeightGrams = 13600')
    lines.append('    double phase1CostUSD = 56252')
    lines.append("")

    # Metadata scope
    lines.append('    def Scope "Metadata"')
    lines.append("    {")
    lines.append('        string generatedBy = "Aquila BOM USD Exporter v1.0"')
    lines.append(f'        string timestamp = "{ts}"')
    lines.append('        string source = "Aquila_Master_BOM_Reconciled_v2 7-5-26.xlsx"')
    lines.append('        string coordinateSystem = "Local frame (mm), Y-up"')
    lines.append('        string[] phases = ["Phase 1 — Operational", "Phase 2 — Firefly Hybrid Mount", "Phase 3 — Bio-Intel DNA Stack"]')
    lines.append("    }")
    lines.append("")

    # Materials scope (shared materials)
    lines.append('    def Scope "Materials"')
    lines.append("    {")
    shared_mats = [
        ("DroneBody_Mat", (0.15, 0.15, 0.18), 0.5, 0.4),
        ("Motor_Mat", (0.6, 0.6, 0.65), 0.8, 0.2),
        ("Prop_Mat", (0.1, 0.1, 0.1), 0.2, 0.3),
    ]
    for name, color, metallic, roughness in shared_mats:
        lines.append(f'        def Material "{name}"')
        lines.append("        {")
        lines.append(f'            token outputs:surface.connect = </AquilaSovereign/Materials/{name}/Shader.outputs:surface>')
        lines.append('            def Shader "Shader"')
        lines.append("            {")
        lines.append('                uniform token info:id = "UsdPreviewSurface"')
        lines.append(f'                color3f inputs:diffuseColor = {fmt_color(color)}')
        lines.append(f'                float inputs:metallic = {metallic}')
        lines.append(f'                float inputs:roughness = {roughness}')
        lines.append('                token outputs:surface')
        lines.append("            }")
        lines.append("        }")
    lines.append("    }")
    lines.append("")

    # Drone platform
    lines.extend(drone_platform_mesh(DRONE_PLATFORM))

    # Phase 1 components
    lines.append('    def Scope "Phase1_Operational"')
    lines.append("    {")
    lines.append('    # Phase 1 — Operational leak-detection stack (fly now)')
    for comp in PHASE1_COMPONENTS:
        safe_name = comp["id"].replace("-", "_")
        if comp["shape"] == "cylinder":
            diam = max(comp["dims"])
            height = comp["dims"][2] if len(comp["dims"]) > 2 else comp["dims"][0]
            lines.extend(cylinder_mesh(
                f"{safe_name}", diam, height,
                comp["pos"], comp["color"],
                bom_id=comp["id"], subsystem=comp["subsystem"], weight=comp["weight"]
            ))
        elif comp["shape"] == "gimbal":
            lines.extend(gimbal_mesh(
                f"{safe_name}", comp["dims"], comp["pos"], comp["color"],
                bom_id=comp["id"], subsystem=comp["subsystem"], weight=comp["weight"]
            ))
        elif comp["shape"] == "frame":
            lines.extend(frame_mesh(
                f"{safe_name}", comp["dims"], comp["pos"], comp["color"],
                bom_id=comp["id"], subsystem=comp["subsystem"], weight=comp["weight"]
            ))
        else:
            lines.extend(box_mesh(
                f"{safe_name}", comp["dims"], comp["pos"], comp["color"],
                bom_id=comp["id"], subsystem=comp["subsystem"], weight=comp["weight"]
            ))
    lines.append("    }")
    lines.append("")

    # Firefly mount (Phase 2)
    lines.append('    def Scope "Phase2_FireflyMount"')
    lines.append("    {")
    lines.append('    # Phase 2 — Firefly hybrid mount components')
    for comp in FIREFLY_MOUNT:
        safe_name = comp["id"].replace("-", "_")
        if comp["shape"] == "cylinder":
            diam = max(comp["dims"])
            height = comp["dims"][2]
            lines.extend(cylinder_mesh(
                f"{safe_name}", diam, height,
                comp["pos"], comp["color"],
                bom_id=comp["id"], subsystem=comp["subsystem"], weight=comp["weight"]
            ))
        elif comp["shape"] == "frame":
            lines.extend(frame_mesh(
                f"{safe_name}", comp["dims"], comp["pos"], comp["color"],
                bom_id=comp["id"], subsystem=comp["subsystem"], weight=comp["weight"]
            ))
        else:
            lines.extend(box_mesh(
                f"{safe_name}", comp["dims"], comp["pos"], comp["color"],
                bom_id=comp["id"], subsystem=comp["subsystem"], weight=comp["weight"]
            ))
    lines.append("    }")
    lines.append("")

    # Bio-Intel stack (Phase 3)
    lines.append('    def Scope "Phase3_BioIntel"')
    lines.append("    {")
    lines.append('    # Phase 3 — Bio-Intel DNA identification stack')
    for comp in BIO_STACK:
        safe_name = comp["id"].replace("-", "_")
        if comp["shape"] == "cylinder":
            diam = max(comp["dims"])
            height = comp["dims"][2]
            lines.extend(cylinder_mesh(
                f"{safe_name}", diam, height,
                comp["pos"], comp["color"],
                bom_id=comp["id"], subsystem=comp["subsystem"], weight=comp["weight"]
            ))
        else:
            lines.extend(box_mesh(
                f"{safe_name}", comp["dims"], comp["pos"], comp["color"],
                bom_id=comp["id"], subsystem=comp["subsystem"], weight=comp["weight"]
            ))
    lines.append("    }")
    lines.append("")

    # Sensor field-of-view visualization (GPR cone, QCL beam)
    lines.append('    def Scope "SensorBeams"')
    lines.append("    {")
    # GPR cone (downward, 30ft depth = ~9m = 9000mm)
    lines.append('        def Mesh "GPR_Cone"')
    lines.append("        {")
    lines.append('            float3[] extent = [(0, 0, 0), (3000, 3000, 9000)]')
    lines.append('            int[] faceVertexCounts = [3, 3, 3, 3, 3, 3, 3, 3]')
    lines.append('            int[] faceVertexIndices = [0,1,2, 0,2,3, 0,3,4, 0,4,5, 0,5,6, 0,6,7, 0,7,8, 0,8,1]')
    lines.append('            point3f[] points = [(0,0,0), (2000,-2000,9000), (-2000,-2000,9000), (-2000,2000,9000), (2000,2000,9000), (3000,0,9000), (0,3000,9000), (-3000,0,9000), (0,-3000,9000)]')
    lines.append('            matrix4d xformOp:transform = ( (1,0,0,0), (0,1,0,0), (0,0,1,0), (0,-120,-40,1) )')
    lines.append('            token[] xformOpOrder = ["xformOp:transform"]')
    lines.append('            rel material:binding = </AquilaSovereign/SensorBeams/GPR_Cone_Mat>')
    lines.append("        }")
    lines.append('        def Material "GPR_Cone_Mat"')
    lines.append("        {")
    lines.append('            token outputs:surface.connect = </AquilaSovereign/SensorBeams/GPR_Cone_Mat/Shader.outputs:surface>')
    lines.append('            def Shader "Shader"')
    lines.append("            {")
    lines.append('                uniform token info:id = "UsdPreviewSurface"')
    lines.append('                color3f inputs:diffuseColor = (0.0, 1.0, 0.5)')
    lines.append('                float inputs:opacity = 0.15')
    lines.append('                token outputs:surface')
    lines.append("            }")
    lines.append("        }")
    lines.append("    }")
    lines.append("")

    # Camera (inspection view)
    lines.append('    def Scope "InspectionCamera"')
    lines.append("    {")
    lines.append('        def Camera "OrbitCamera"')
    lines.append("        {")
    lines.append('            float focalLength = 24')
    lines.append('            float horizontalAperture = 36')
    lines.append('            float verticalAperture = 24')
    lines.append('            float2 clippingRange = (10, 100000)')
    lines.append('            double3 xformOp:translate = (2500, 1500, 2500)')
    lines.append('            double3 xformOp:rotateXYZ = (-20, 35, 0)')
    lines.append('            token[] xformOpOrder = ["xformOp:translate", "xformOp:rotateXYZ"]')
    # Animate orbit
    lines.append('            double3 xformOp:translate.timeSamples = {')
    for frame in range(1, 901, 30):
        ang = (frame / 900) * 2 * math.pi
        r = 3000
        x = r * math.cos(ang)
        z = r * math.sin(ang)
        y = 1500 + 500 * math.sin(frame / 30.0 * 0.3)
        lines.append(f'                {frame}: ({x:.2f}, {y:.2f}, {z:.2f}),')
    lines.append('            }')
    lines.append("        }")
    lines.append("    }")
    lines.append("")

    lines.append("}")
    return "\n".join(lines)


def main():
    parser = argparse.ArgumentParser(description="AQUILA Sovereign Platform USD Exporter")
    parser.add_argument("--output", "-o", default="aquila_sovereign_platform.usda", help="Output USDA file")
    parser.add_argument("--z-up", action="store_true", help="Use Z-up axis (Isaac Sim convention)")
    args = parser.parse_args()

    up_axis = "Z" if args.z_up else "Y"
    usda = build_usda(up_axis)

    out_path = Path(args.output)
    out_path.write_text(usda, encoding="utf-8")
    size_kb = out_path.stat().st_size / 1024

    total = len(PHASE1_COMPONENTS) + len(FIREFLY_MOUNT) + len(BIO_STACK)
    print(f"[Aquila USD] Wrote {out_path} ({size_kb:.1f} KB)")
    print(f"[Aquila USD] Platform: Firefly hexacopter ({DRONE_PLATFORM['arms']} arms, {DRONE_PLATFORM['prop_diam_mm']}mm props)")
    print(f"[Aquila USD] Components: {total} BOM items across 3 phases")
    print(f"  - Phase 1 (Operational): {len(PHASE1_COMPONENTS)} components")
    print(f"  - Phase 2 (Firefly Mount): {len(FIREFLY_MOUNT)} components")
    print(f"  - Phase 3 (Bio-Intel): {len(BIO_STACK)} components")
    print(f"[Aquila USD] Animated: 6 propellers (720 deg/sec), orbit camera (900 frames)")
    print(f"[Aquila USD] Up axis: {up_axis}")
    print(f"[Aquila USD] Compatible with NVIDIA Omniverse Kit, Isaac Sim, USDView")


if __name__ == "__main__":
    main()
