# Autochtone Boomsoorten en Struiken (ABS) to Darwin Core Mapping

This repository contains the workflow and scripts for mapping the **Autochtone Boomsoorten (ABS)** Microsoft Access database to the **Darwin Core (DwC)** standard.

The resulting dataset is optimized for publication on the **Vlaams Biodiversiteitsportaal (VBP)**.

## Project Overview

The project involves extracting survey and species data from an legacy MS Access database (`.mdb`), performing coordinate transformations (Lambert 72 to WGS84), and aligning native status codes with international biodiversity data standards.

### Key Features
- **Darwin Core Compliant**: Maps database fields to standard DwC terms like `occurrenceID`, `eventDate`, `scientificName`, and `establishmentMeans`.
- **Coordinate Transformation**: Precise conversion of Belgian Lambert 72 (EPSG:31370) to WGS84 (EPSG:4326) with 5-decimal precision.
- **Date Extraction**: Advanced parsing of specific survey dates from internal numbering systems.
- **Status Mapping**: Translation of "autochtone quotering" codes to standardized `establishmentMeans` and descriptive `occurrenceRemarks`.

## Repository Structure

- `src/map_to_dwc_fast.ps1`: Core PowerShell script for high-performance mapping and CSV generation.
- `src/abs_mapping_workflow.Rmd`: R Markdown project documenting the entire process with tidyverse and interactive mapping.
- `map.html`: Interactive Leaflet map for visualizing the 70,000+ survey locations.
- `data/event_core.csv`: The primary Darwin Core Event dataset.
- `data/occurrence_extension.csv`: The Occurrence extension linking species to events.
- `data/measurement_or_fact.csv`: The MeasurementOrFact extension for tree attributes and site measurements.
- `data/map_data.js`: Clustered data for the interactive map.
## Usage

### PowerShell
Run the following command to regenerate the Darwin Core CSV:
```powershell
powershell -ExecutionPolicy Bypass -File src/map_to_dwc_fast.ps1
```

### R Markdown
Open `src/abs_mapping_workflow.Rmd` in RStudio and click **Knit** to generate a full report with an interactive map.

## License

The data is published under the [CC0 1.0 Universal (CC0 1.0) Public Domain Dedication](http://creativecommons.org/publicdomain/zero/1.0/).

---
*Created as part of the Vlaams Biodiversiteitsportaal data standardization initiative.*
