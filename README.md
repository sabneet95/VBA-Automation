# VBA Automation for Atmospheric Sciences and Neurophysics

This repository contains VBA-based automation solutions designed specifically for the atmospheric sciences and neurophysics domains. It includes robust macros for processing climate and weather data as well as neuroimaging data—featuring advanced calculations, dynamic chart generation, and automated formatting.

![VBA Image](https://github.com/sabneet95/VBA-Automation/blob/master/vba.jpg)

→ **Please note:** The provided code is domain‑specific and may require modification to suit other projects.

---

## Table of Contents

- [Overview](#overview)
- [Key Features](#key-features)
- [Directory Structure](#directory-structure)
- [Requirements](#requirements)
- [Installation & Getting Started](#installation--getting-started)
- [Usage](#usage)
  - [Climate and Weather Macros](#climate-and-weather-macros)
  - [Neuro Macros](#neuro-macros)
- [Contributing](#contributing)
- [License](#license)
- [Future Work](#future-work)

---

## Overview

This repository is a collection of VBA macros tailored for two specialized domains:

**Atmospheric Sciences:**  
Automates the retrieval, processing, and visualization of climate and weather data. The macros dynamically insert calculation columns, build formulas for seasonal and annual averages, and generate charts with trendlines.

**Neurophysics:**  
Provides automation tools for processing neuroimaging data—including SUV and SUVr calculations, frame sequence fixes, ROI processing, and comprehensive graph generation. The code also upgrades legacy files and customizes outputs with interactive prompts.

These tools are intended for researchers and professionals who require robust Excel‑based automation tailored to their specialized data.

---

## Key Features

- **Domain‑Specific Automation:**  
  - Climate and Weather modules automatically retrieve and process atmospheric data.
  - The Neuro module rearranges frame sequences, extracts numeric identifiers, and processes ROI data for neuroimaging studies.

- **Dynamic Data Handling:**  
  - Automatic detection of data ranges, dynamic column insertion, and formula creation.
  - Robust error handling and restoration of Excel settings to maintain performance and stability.

- **Advanced Charting:**  
  - Multiple charts are generated automatically with trendlines, customized axis settings, and legend support.
  - Graphs are arranged and resized dynamically for optimal presentation.

- **User Interaction:**  
  - Prompts for file upgrades, frame ordering, weight/dose input, and time interval customization via user forms.

- **Well-Documented Code:**  
  - Each module follows modern VBA standards with explicit variable declarations, detailed inline comments, and comprehensive header documentation.

---

## Directory Structure

- **Atmospheric_Sciences/**
  - **climate.bas** (Macro for climate data automation)
  - **weather.bas** (Macro for weather data processing and charting)
- **Neurophysics/**
  - **neuro.bas** (Macro for neuroimaging data processing and graph generation)
- **vba.jpg** (Representative image of the VBA project)
- **README.md** (This documentation file)

---

## Requirements

- **VBA 7 or higher**  
  See Getting Started with VBA in Office at https://docs.microsoft.com/en-us/office/vba/library-reference/concepts/getting-started-with-vba-in-office

- **Microsoft Excel 2016 or later** (Tested on Excel 2016 64‑bit)

---

## Installation & Getting Started

1. **Clone the Repository:**  
   Run
   ```bash
   git clone https://github.com/sabneet95/VBA-Automation.git
   ```
   and then change directory with

   ```bash
   cd VBA-Automation
   ```

3. **Open the Project in Excel:**  
   Open your Excel workbook and import the `.bas` modules from the *Atmospheric_Sciences* and *Neurophysics* folders into your VBA project.

4. **Enable the Developer Tab:**  
   In Excel, ensure the Developer tab is visible to access the VBA editor and run the macros.

---

## Usage

### Climate and Weather Macros

- **Climate Macro:**  
  Processes climate data by retrieving web-based data and formatting it with calculated temperature averages.

- **Weather Macro:**  
  Automates the processing of weather data, inserting columns for annual and seasonal averages (both temperature and precipitation) and generating multiple charts with trendlines.

To run these macros:  
1. Open your Excel workbook.  
2. Go to the Developer tab and click on "Macros."  
3. Select either `Climate` or `Weather` and click "Run."  

The macros will create new worksheets, insert necessary formulas, and generate charts automatically.

---

### Neuro Macros

The Neuro macro is designed for neuroimaging data automation. It performs several key functions:

- **File Upgrade:** Upgrades non-XML workbooks automatically.
- **Frame Sequence Fix:** Rearranges frame sequences and extracts numeric identifiers from ROI names.
- **ROI Processing:** Converts ROI labels, calculates averages, and builds SUV/SUVr formulas.
- **Graph Generation:** Automatically creates multiple charts and multiplot graphs with customized formatting.
- **User Inputs:** Prompts for weight, dose, and time interval adjustments via interactive user forms.

To run the Neuro macro:  
1. Open your Excel workbook containing neuroimaging data.  
2. Go to the Developer tab, select the `Neuro` macro, and click "Run."  
3. Follow the on-screen prompts (e.g., frame ordering, weight/dose input).

---

## Contributing

Contributions are welcome! To contribute:  
1. Open an issue to discuss your proposed changes.  
2. Ensure your code follows modern VBA practices (Option Explicit, proper error handling, and thorough inline comments).  
3. Submit a pull request with clear descriptions and include tests or documentation as needed.

---

## License

This repository is licensed under the [MIT License](https://choosealicense.com/licenses/mit/).

---

## Future Work

Planned improvements include:  
- Enhanced error logging and debugging support.  
- Additional modules for advanced statistical analysis.  
- More interactive user forms for parameter customization.  
- Expanded visualization options and automated report generation.  
- Performance optimizations and broader compatibility across Excel versions.

---

*For questions or further information, please check the Issues section or contact the maintainers.*
