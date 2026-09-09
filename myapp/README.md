# STATCAL ONLINE Likert Scale Distribution and Stacked Bar Chart Visualizer

This R Shiny application analyzes Likert-scale survey data with flexible question item selection, response-label ordering, frequency and percentage tables, split tables by a categorical variable, stacked bar charts, PNG export, and Excel export.

## Required packages

```r
install.packages(c(
  "shiny", "shinydashboard", "DT", "readxl", "dplyr", "tidyr",
  "ggplot2", "shinycssloaders", "scales", "openxlsx"
))
```

## Run the application

```r
shiny::runApp(".")
```

## Included files

- `app.R`
- `data_likert.xlsx`
- `README.md`

## Main features

- Flexible item question selection.
- Flexible Likert response order, for example: `STS, TS, N, S, SS`.
- Frequency and percentage tables.
- Split table by one categorical variable such as `Jenis Kelamin`.
- Stacked bar charts with frequency, percentage, or combined labels.
- Vertical and horizontal chart orientation.
- Manual bar colors using HEX codes.
- Several publication-style chart backgrounds.
- High-resolution PNG export with default 1200 DPI.
- Excel export for raw data, frequency table, split frequency table, chart data, and metadata.
