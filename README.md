# AQUILA Sovereign Platform

**Multi-mineral detection, eDNA capture, and triple-verified target validation drone.**

## Overview

The AQUILA Sovereign platform is an integrated mineral-detection and orphan-well validation drone designed for the Osage Minerals Council's Arbuckle Redevelopment mission. It operates across two deployment phases:

- **Phase 1:** Freefly Alta X electric platform for parking-lot and short-range validation
- **Phase 2:** Parallel Flight Firefly gas-hybrid platform for full 100-lb payload and 1.4-hour endurance

The on-board forensic vision logic detects mineral deposits, wellhead erosion, and bio-cracking sweet spots simultaneously using a triple-verified RLVR protocol.

## Repository Contents

### Bill of Materials
- `Aquila_Master_BOM_Reconciled_v2 7-5-26 (version 1).xlsx` — 65 reconciled BOM line items with dimensions, weights, costs, and subsystem assignments

### USD Digital Twin
- `Aquila_Sovereign_Complete.usda` — Complete OpenUSD stage (Pixar 0.26.8) with 144 mesh prims from BOM + 57 engineering spec scopes
- `aquila_sovereign_platform.usda` — Base platform USD stage
- `convert_all_to_usd.py` — Converter script (BOM xlsx + .docx specs → USD)
- `export_aquila_usd.py` — Base USDA exporter from BOM data

### Engineering Specifications (57 .docx files)
Engineering design innovations covering:
- **Mechanical:** O-ring retainers, piston ring seals, helical inserts, expansion chambers
- **Electromagnetic:** Clutch torque decoupling, dual-motor segmented torque
- **Vibration:** Conical rubber isolators, quasi-zero stiffness isolation, passive-active hybrid suppression
- **Thermal:** Phase change material buffers, condensation pipes, heat pipe thermal highways
- **Filtration:** Negative pressure enclosures, staged cyclone HEPA
- **EMI Shielding:** Graphene-metal oxide composite films, multi-layer conductive coatings, segmented ground planes, ferrite core cable resonance
- **Bio-Protection:** DrACO-inspired autonomous sampling subsystem

### Autonomous Scientific Sampling Subsystem
- `Autonomous Scientific Sampling Subsystem/` — DrACO-inspired sampling module documentation

## Sensor Stack

| Subsystem | Component | Function |
|-----------|-----------|----------|
| Compute Core | NVIDIA Jetson AGX Orin 64 GB | 275 TOPS edge AI |
| Radar | AERIS-10 + AQUILA | Phased-array surface tracking |
| Subsurface Radar | Zond Aero LF + Warp RTM | Karst void mapping |
| Chemical — QCL | Hamamatsu L12004-2190H-C | 4.6 µm mid-IR methane detection |
| Chemical — Detector | Vigo PVI-4TE-5 | 3–5 µm QCL backscatter |
| Photoacoustic | Brüel & Kjær Type 4955 | 15 dBA sensitivity |
| EO/IR Imaging | FLIR Hadron 640R | LWIR + RGB gimballed |
| Photogrammetry | Sony ILX-LR1 (61 MP) | Full-frame ortho mapping |
| DNA Sequencer | Oxford Nanopore MinION Mk1C | Real-time eDNA surveillance |
| Sample Acquisition | STaPLE + PlanetVac + DrACO | Brine/soil for Li/Mg assay |
| Multimodal AI | Nemotron-3 Nano Omni | Real-time ppm-vs-price reasoning |

## Triple-Verified Target Protocol

| Layer | Sensor | Trigger | FDR Target |
|-------|--------|---------|------------|
| Physical | Zond Aero LF + Warp RTM | Hyperbola collapse | < 10% |
| Chemical | Hamamatsu QCL + spec | Spectral peak at 4.6 µm | < 10% |
| Biological | Oxford Nanopore MinION | Sb-resistant bacteria DNA | < 5% |
| **Triple-Verified** | All three layers | All triggers fire | **< 2%** |

## Mission Context

Part of the **Sovereign Mission** for the Osage Minerals Council — Arbuckle Redevelopment, Phases 1–20.1. The drone validates orphan-well clusters and bio-cracking candidates before any plugging capital is committed.

## License

Sovereign Confidential — Council Eyes Only

## Author

Dr. White Sutherland — Sovereign Mission Lead
Aquila Geological Systems
