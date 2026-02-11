# MetraCAD v1.0.0
### CAD LISP Plugin for Automated Length Quantification

MetraCAD is a high-performance LISP utility designed for AutoCAD and ZWCAD environments. It streamlines the process of measuring cumulative lengths across multiple entity types with automated reporting and data export capabilities.

## Technical Specifications

| Feature | Details |
| :--- | :--- |
| **Supported Entities** | Line, Polyline (2D/3D), Arc, Circle, Spline, Ellipse |
| **Selection Logic** | Manual, Layer-based, or Color Index (ACI) |
| **Precision** | Configurable decimal places (2, 3, or 4) |
| **Environments** | AutoCAD 2010+, ZWCAD 2020+ |
| **Dependencies** | Visual LISP (VLAX) Extension |

## Command Reference

| Command | Function |
| :--- | :--- |
| `METRACAD` | Initialize primary measurement routine and reporting. |
| `METRACADEXPORT` | Export session data to TXT or CSV (Excel compatible). |
| `METRACADHISTORY` | Display cumulative session measurement log. |
| `METRACADHELP` | Access internal documentation and syntax. |

## Installation

### A. Volatile Load (Current Session)
1. Execute `APPLOAD` in the command line.
2. Select `metracad.lsp` and click **Load**.

### B. Persistent Load (Startup Suite)
1. Execute `APPLOAD`.
2. Navigate to **Startup Suite** > **Contents**.
3. Add `metracad.lsp` to the registry.

## Data Structure Example (Output)

MetraCAD generates a structured report within the command line and an optional system alert:

* **Header**: Filename, timestamp, and active units.
* **Type Breakdown**: Count and sum per entity class.
* **Layer Breakdown**: Detailed analysis if objects span multiple layers.

## Development Roadmap

* **v1.x**: Multi-layer selection logic and Area (m²) computation.
* **v2.0**: Migration to .NET API (C#) for multithreading and WPF UI.

## License
MIT License. Copyright (c) 2026 Jean Carlos.