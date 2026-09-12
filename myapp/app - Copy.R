# ============================================================
# STATCAL ONLINE - Likert Scale Distribution and Stacked Bar Chart Visualizer
# R Shiny Version
# ============================================================
# Required packages:
# install.packages(c(
#   "shiny", "shinydashboard", "DT", "readxl", "dplyr", "tidyr",
#   "ggplot2", "shinycssloaders", "scales", "openxlsx",
#   "colourpicker", "officer", "flextable", "base64enc"
# ))

required_packages <- c(
  "shiny", "shinydashboard", "DT", "readxl", "dplyr", "tidyr",
  "ggplot2", "shinycssloaders", "scales", "openxlsx",
  "colourpicker", "officer", "flextable", "base64enc"
)

missing_packages <- required_packages[!vapply(required_packages, requireNamespace, logical(1), quietly = TRUE)]
if (length(missing_packages) > 0) {
  stop(
    "Please install the following R packages first: ",
    paste(missing_packages, collapse = ", "),
    call. = FALSE
  )
}

library(shiny)
library(shinydashboard)
library(DT)
library(readxl)
library(dplyr)
library(tidyr)
library(ggplot2)
library(shinycssloaders)
library(scales)
library(openxlsx)
library(colourpicker)
library(officer)
library(flextable)


# ============================================================
# SHINYLIVE / STATIC HOSTING DOWNLOAD STRATEGY
# ============================================================
# Binary exports are NOT delivered through Shiny downloadHandler URLs.
# In GitHub Pages + Shinylive/webR, those dynamic session URLs can fail
# when the Service Worker loses the parent-page/session mapping.
#
# Instead, STATCAL:
# 1. creates the file inside webR's virtual filesystem,
# 2. validates the generated binary file,
# 3. Base64-encodes it in R,
# 4. sends it to JavaScript with session$sendCustomMessage(),
# 5. reconstructs it as a browser Blob and starts a normal download.
#
# This approach is used for PNG, Excel, and Word exports.


# ============================================================
# CONSTANTS
# ============================================================

APP_NAME <- "STATCAL ONLINE"
APP_TITLE <- "Likert Scale Distribution and Stacked Bar Chart Visualizer"
APP_UPDATED <- "Last updated on September 12, 2026"
SAMPLE_DATA_PATH <- "data_likert.xlsx"
WEBSITE_URL <- "https://statcal.com/"
URL_DATA_TRAINING <- "https://drive.google.com/drive/folders/1QMSiOx2NeB6Rj1j9lmmEK_Im-8YaLFNE?usp=sharing"
STATCAL_ONLINE_URL <- "https://statcal.com/statcal%20online.html"
URL_DATASET <- URL_DATA_TRAINING
DEFAULT_LIKERT_ORDER <- "STS, TS, N, S, SS"

THEMES <- list(
  "White Publication" = list(
    figure_facecolor = "white", axes_facecolor = "white", text_color = "#111111",
    grid_color = "#D9D9D9", spine_color = "#222222"
  ),
  "Light Gray Editorial" = list(
    figure_facecolor = "#F7F7F7", axes_facecolor = "#FFFFFF", text_color = "#111111",
    grid_color = "#D0D0D0", spine_color = "#333333"
  ),
  "Warm Ivory Journal" = list(
    figure_facecolor = "#FBF7EF", axes_facecolor = "#FFFDF8", text_color = "#1F1F1F",
    grid_color = "#DDD4C4", spine_color = "#3A3A3A"
  ),
  "Cool Blue Scientific" = list(
    figure_facecolor = "#F3F7FB", axes_facecolor = "#FFFFFF", text_color = "#0B1F33",
    grid_color = "#C8D6E5", spine_color = "#1F4E79"
  ),
  "Dark Navy Presentation" = list(
    figure_facecolor = "#0B1320", axes_facecolor = "#111C2E", text_color = "#FFFFFF",
    grid_color = "#3B4A5F", spine_color = "#B8C7D9"
  ),
  "Minimal Scopus Style" = list(
    figure_facecolor = "#FFFFFF", axes_facecolor = "#FFFFFF", text_color = "#111111",
    grid_color = "#EAEAEA", spine_color = "#111111"
  ),
  "Soft Blue Journal" = list(
    figure_facecolor = "#F4F8FC", axes_facecolor = "#FFFFFF", text_color = "#102A43",
    grid_color = "#D8E6F2", spine_color = "#243B53"
  )
)

DEFAULT_COLORS <- c(
  "#B2182B", "#EF8A62", "#F7F7F7", "#67A9CF", "#2166AC",
  "#1B7837", "#762A83", "#8C510A", "#4D4D4D", "#F46D43"
)


LIKERT_PALETTES <- list(
  "Scopus-inspired Orange–Blue" = c("#E97132", "#F4B183", "#E7E6E6", "#5B9BD5", "#1F4E79"),
  "Scopus-inspired Orange–Teal" = c("#E97132", "#F4B183", "#F2F2F2", "#70C1B3", "#0B7189"),
  "Scopus-inspired Orange–Navy" = c("#F28E2B", "#FFBE7D", "#BAB0AC", "#76B7B2", "#4E79A7"),
  "Publication Red–Blue" = c("#B2182B", "#EF8A62", "#F7F7F7", "#67A9CF", "#2166AC"),
  "Publication Brown–Teal" = c("#8C510A", "#D8B365", "#F5F5F5", "#5AB4AC", "#01665E"),
  "Colorblind Friendly" = c("#D55E00", "#E69F00", "#F0E442", "#56B4E9", "#0072B2"),
  "Viridis-like" = c("#440154", "#3B528B", "#21918C", "#5EC962", "#FDE725"),
  "Soft Pastel" = c("#D98880", "#F5B7B1", "#F7F7F7", "#AED6F1", "#5DADE2"),
  "Monochrome Blue" = c("#D6EAF8", "#AED6F1", "#85C1E9", "#5DADE2", "#2E86C1"),
  "Manual / Custom" = NULL
)

DEFAULT_LEGEND_LABELS <- c(
  "STS" = "Sangat Tidak Setuju",
  "TS" = "Tidak Setuju",
  "N" = "Netral",
  "S" = "Setuju",
  "SS" = "Sangat Setuju"
)

# ============================================================
# HELPER FUNCTIONS
# ============================================================

clean_dataframe <- function(df) {
  names(df) <- trimws(gsub("\\s+", " ", as.character(names(df))))
  df <- df[rowSums(is.na(df)) < ncol(df), , drop = FALSE]
  unnamed_cols <- grepl("^unnamed", tolower(names(df)))
  if (any(unnamed_cols)) {
    keep_unnamed <- vapply(df[unnamed_cols], function(x) !all(is.na(x)), logical(1))
    drop_names <- names(df)[unnamed_cols][!keep_unnamed]
    if (length(drop_names) > 0) df <- df[, !names(df) %in% drop_names, drop = FALSE]
  }
  rownames(df) <- NULL
  as.data.frame(df)
}

make_display_safe <- function(df) {
  df <- as.data.frame(df)
  for (nm in names(df)) {
    if (is.factor(df[[nm]])) df[[nm]] <- as.character(df[[nm]])
  }
  df
}

sorted_unique_values <- function(x) {
  vals <- unique(x[!is.na(x)])
  vals[order(as.character(vals))]
}

get_theme <- function(theme_name) {
  if (is.null(theme_name) || length(theme_name) == 0 || is.na(theme_name) || !(theme_name %in% names(THEMES))) {
    return(THEMES[["White Publication"]])
  }
  THEMES[[theme_name]]
}

safe_theme_bg <- function(theme_name) {
  tryCatch(get_theme(theme_name)$figure_facecolor, error = function(e) "white")
}

safe_number <- function(x, default_value, min_value = NULL, max_value = NULL) {
  out <- suppressWarnings(as.numeric(x))
  if (length(out) == 0 || is.na(out) || !is.finite(out)) out <- default_value
  if (!is.null(min_value)) out <- max(out, min_value)
  if (!is.null(max_value)) out <- min(out, max_value)
  out
}

parse_likert_levels <- function(x) {
  if (is.null(x) || length(x) == 0 || is.na(x)) x <- DEFAULT_LIKERT_ORDER
  levels <- trimws(unlist(strsplit(as.character(x), ",", fixed = TRUE)))
  levels <- levels[nzchar(levels)]
  levels <- unique(levels)
  if (length(levels) == 0) levels <- c("STS", "TS", "N", "S", "SS")
  levels
}

parse_split_order <- function(x, available_values) {
  available_values <- as.character(available_values)
  available_values <- available_values[!is.na(available_values) & nzchar(available_values)]
  available_values <- unique(available_values)
  if (length(available_values) == 0) return(character(0))
  if (is.null(x) || length(x) == 0 || is.na(x) || !nzchar(trimws(as.character(x)))) {
    return(available_values)
  }
  wanted <- trimws(unlist(strsplit(as.character(x), ",", fixed = TRUE)))
  wanted <- wanted[nzchar(wanted)]
  ordered <- wanted[wanted %in% available_values]
  remaining <- setdiff(available_values, ordered)
  unique(c(ordered, remaining))
}

safe_token <- function(x) {
  x <- gsub("[^A-Za-z0-9_]", "_", as.character(x))
  x <- gsub("_+", "_", x)
  x
}

color_input_id <- function(x) paste0("color_", safe_token(x))
legend_input_id <- function(x) paste0("legend_label_", safe_token(x))

is_hex_color <- function(x) {
  is.character(x) && length(x) == 1 && grepl("^#([0-9A-Fa-f]{6}|[0-9A-Fa-f]{8})$", x)
}

palette_colors <- function(palette_name, n, reverse = FALSE) {
  n <- max(1L, as.integer(n))
  base_cols <- LIKERT_PALETTES[[palette_name]]
  if (is.null(base_cols) || length(base_cols) == 0) base_cols <- DEFAULT_COLORS
  if (length(base_cols) == n) {
    out <- base_cols
  } else if (n == 1) {
    out <- base_cols[ceiling(length(base_cols) / 2)]
  } else {
    out <- grDevices::colorRampPalette(base_cols)(n)
  }
  if (isTRUE(reverse)) out <- rev(out)
  out
}

get_manual_colors <- function(levels, input) {
  out <- character(length(levels))
  defaults <- palette_colors("Scopus-inspired Orange–Blue", length(levels))
  for (i in seq_along(levels)) {
    val <- input[[color_input_id(levels[i])]]
    if (is.null(val) || length(val) == 0 || !is_hex_color(val)) val <- defaults[i]
    out[i] <- val
  }
  names(out) <- levels
  out
}

get_chart_colors <- function(levels, input) {
  palette_name <- input$chart_palette
  if (is.null(palette_name) || !(palette_name %in% names(LIKERT_PALETTES))) {
    palette_name <- "Scopus-inspired Orange–Blue"
  }
  if (palette_name == "Manual / Custom") {
    return(get_manual_colors(levels, input))
  }
  out <- palette_colors(palette_name, length(levels), isTRUE(input$chart_reverse_palette))
  names(out) <- levels
  out
}

default_legend_label <- function(level) {
  if (level %in% names(DEFAULT_LEGEND_LABELS)) DEFAULT_LEGEND_LABELS[[level]] else as.character(level)
}

get_legend_labels <- function(levels, input) {
  out <- vapply(levels, function(lv) {
    val <- input[[legend_input_id(lv)]]
    if (is.null(val) || length(val) == 0 || is.na(val) || !nzchar(trimws(as.character(val)))) {
      default_legend_label(lv)
    } else {
      trimws(as.character(val))
    }
  }, character(1))
  names(out) <- levels
  out
}

detect_likert_item_columns <- function(df, levels) {
  levels <- as.character(levels)
  if (length(levels) == 0) return(character(0))
  cols <- names(df)
  good <- vapply(df, function(x) {
    z <- trimws(as.character(x))
    z <- z[!is.na(z) & nzchar(z)]
    if (length(z) == 0) return(FALSE)
    mean(z %in% levels) >= 0.90 && length(unique(z)) <= max(length(levels) + 1L, 7L)
  }, logical(1))
  out <- cols[good]
  if (length(out) == 0) {
    out <- cols[grepl("^[A-Za-z]+[0-9]+$", cols)]
  }
  if (length(out) == 0) {
    out <- grep("pertanyaan|question|item", cols, ignore.case = TRUE, value = TRUE)
  }
  out
}

statcal_theme_gg <- function(theme_name,
                             title_size = 16,
                             subtitle_size = 11,
                             axis_title_size = 11,
                             axis_text_size = 9,
                             legend_title_size = 10,
                             legend_text_size = 9,
                             legend_position = "Right",
                             x_text_angle = 0) {
  th <- get_theme(theme_name)
  legend_pos <- ifelse(is.null(legend_position) || legend_position == "None / Hide legend", "none", tolower(legend_position))
  theme_minimal(base_size = axis_text_size) +
    theme(
      plot.background = element_rect(fill = th$figure_facecolor, color = NA),
      panel.background = element_rect(fill = th$axes_facecolor, color = NA),
      panel.grid.major = element_line(color = th$grid_color, linewidth = 0.35),
      panel.grid.minor = element_line(color = th$grid_color, linewidth = 0.15),
      axis.text = element_text(color = th$text_color, size = axis_text_size),
      axis.text.x = element_text(angle = x_text_angle, hjust = ifelse(x_text_angle == 0, 0.5, 1), color = th$text_color, size = axis_text_size),
      axis.title = element_text(color = th$text_color, face = "bold", size = axis_title_size),
      plot.title = element_text(color = th$text_color, face = "bold", size = title_size),
      plot.subtitle = element_text(color = th$text_color, size = subtitle_size),
      legend.text = element_text(color = th$text_color, size = legend_text_size),
      legend.title = element_text(color = th$text_color, face = "bold", size = legend_title_size),
      legend.background = element_rect(fill = th$figure_facecolor, color = NA),
      legend.position = legend_pos
    )
}

# ============================================================
# LIKERT FREQUENCY TABLES
# ============================================================
# Important methodological note:
# The frequency tables and the stacked bar chart must be calculated from the
# same long-format distribution object. This prevents inconsistent results
# between the Split Table tab and the Stacked Bar Chart tab.

create_empty_likert_long_data <- function(item_cols, levels, split_levels = "All") {
  out <- expand.grid(
    `Item Pertanyaan` = item_cols,
    Split = split_levels,
    Response = levels,
    stringsAsFactors = FALSE,
    check.names = FALSE
  )
  out$Frequency <- 0
  out$Total <- 0
  out$Percentage <- NA_real_
  out$`Item Pertanyaan` <- factor(out$`Item Pertanyaan`, levels = item_cols)
  out$Split <- factor(out$Split, levels = split_levels)
  out$Response <- factor(out$Response, levels = levels)
  as.data.frame(out)
}

build_likert_long_data <- function(df, item_cols, levels, split_col = "None", digits = 2, split_order = NULL) {
  validate(need(length(item_cols) > 0, "Please select at least one item question."))
  item_cols <- item_cols[item_cols %in% names(df)]
  validate(need(length(item_cols) > 0, "The selected item question variables were not found in the dataset."))
  
  use_split <- !is.null(split_col) && split_col != "None" && split_col %in% names(df)
  cols <- item_cols
  if (use_split) cols <- c(split_col, item_cols)
  
  tmp <- df[, cols, drop = FALSE]
  
  if (use_split) {
    names(tmp)[names(tmp) == split_col] <- ".Split"
    tmp$.Split <- trimws(as.character(tmp$.Split))
    available_split_values <- sorted_unique_values(tmp$.Split)
    split_levels <- parse_split_order(split_order, available_split_values)
  } else {
    tmp$.Split <- "All"
    split_levels <- "All"
  }
  
  if (length(split_levels) == 0) {
    return(create_empty_likert_long_data(item_cols, levels, "All"))
  }
  
  long_raw <- tmp %>%
    tidyr::pivot_longer(cols = dplyr::all_of(item_cols), names_to = "Item Pertanyaan", values_to = "Response") %>%
    mutate(
      `Item Pertanyaan` = as.character(`Item Pertanyaan`),
      Response = trimws(as.character(Response)),
      Split = trimws(as.character(.Split))
    ) %>%
    filter(
      !is.na(`Item Pertanyaan`), `Item Pertanyaan` %in% item_cols,
      !is.na(Response), Response %in% levels,
      !is.na(Split), Split %in% split_levels
    ) %>%
    mutate(
      `Item Pertanyaan` = factor(`Item Pertanyaan`, levels = item_cols),
      Split = factor(Split, levels = split_levels),
      Response = factor(Response, levels = levels)
    )
  
  if (nrow(long_raw) == 0) {
    return(create_empty_likert_long_data(item_cols, levels, split_levels))
  }
  
  long_df <- long_raw %>%
    count(`Item Pertanyaan`, Split, Response, name = "Frequency", .drop = FALSE) %>%
    tidyr::complete(
      `Item Pertanyaan` = factor(item_cols, levels = item_cols),
      Split = factor(split_levels, levels = split_levels),
      Response = factor(levels, levels = levels),
      fill = list(Frequency = 0)
    ) %>%
    group_by(`Item Pertanyaan`, Split) %>%
    mutate(
      Total = sum(Frequency, na.rm = TRUE),
      Percentage = ifelse(Total > 0, Frequency / Total * 100, NA_real_)
    ) %>%
    ungroup()
  
  long_df$Percentage <- round(long_df$Percentage, digits)
  long_df$`Item Pertanyaan` <- factor(as.character(long_df$`Item Pertanyaan`), levels = item_cols)
  long_df$Split <- factor(as.character(long_df$Split), levels = split_levels)
  long_df$Response <- factor(as.character(long_df$Response), levels = levels)
  as.data.frame(long_df)
}

format_likert_wide_table_from_long <- function(long_df, levels, digits = 2, split_col = "None", include_split = FALSE) {
  validate(need(nrow(long_df) > 0, "No Likert distribution data are available."))
  
  item_order <- base::levels(long_df$`Item Pertanyaan`)
  if (is.null(item_order) || length(item_order) == 0) item_order <- unique(as.character(long_df$`Item Pertanyaan`))
  split_order <- base::levels(long_df$Split)
  if (is.null(split_order) || length(split_order) == 0) split_order <- unique(as.character(long_df$Split))
  
  rows <- list()
  for (sp in split_order) {
    for (item in item_order) {
      sub <- long_df[as.character(long_df$Split) == sp & as.character(long_df$`Item Pertanyaan`) == item, , drop = FALSE]
      
      if (include_split && !(is.null(split_col) || split_col == "None")) {
        row <- data.frame(
          `Split Variable` = split_col,
          `Split Category` = sp,
          `Item Pertanyaan` = item,
          check.names = FALSE
        )
      } else {
        row <- data.frame(`Item Pertanyaan` = item, check.names = FALSE)
      }
      
      total_value <- if (nrow(sub) > 0) max(sub$Total, na.rm = TRUE) else 0
      if (!is.finite(total_value)) total_value <- 0
      
      for (lv in levels) {
        lv_sub <- sub[as.character(sub$Response) == lv, , drop = FALSE]
        freq_value <- if (nrow(lv_sub) > 0) lv_sub$Frequency[1] else 0
        pct_value <- if (nrow(lv_sub) > 0) lv_sub$Percentage[1] else if (total_value > 0) 0 else NA_real_
        row[[paste0(lv, " f")]] <- as.numeric(freq_value)
        row[[paste0(lv, " %")]] <- round(as.numeric(pct_value), digits)
      }
      row[["Total f"]] <- as.numeric(total_value)
      row[["Total %"]] <- ifelse(total_value > 0, round(100, digits), NA_real_)
      rows[[length(rows) + 1]] <- row
    }
  }
  
  dplyr::bind_rows(rows)
}

compute_likert_frequency_table <- function(df, item_cols, levels, digits = 2) {
  long_df <- build_likert_long_data(df, item_cols, levels, split_col = "None", digits = digits)
  format_likert_wide_table_from_long(long_df, levels, digits, include_split = FALSE)
}

compute_likert_split_table <- function(df, item_cols, levels, split_col = "None", digits = 2, split_order = NULL) {
  if (is.null(split_col) || split_col == "None" || !(split_col %in% names(df))) {
    return(compute_likert_frequency_table(df, item_cols, levels, digits))
  }
  long_df <- build_likert_long_data(df, item_cols, levels, split_col = split_col, digits = digits, split_order = split_order)
  format_likert_wide_table_from_long(long_df, levels, digits, split_col = split_col, include_split = TRUE)
}

# ============================================================
# STACKED BAR CHART
# ============================================================

create_stacked_bar_plot <- function(chart_df, levels, manual_colors,
                                    legend_labels = NULL,
                                    legend_title = "Likert Response",
                                    title, subtitle,
                                    x_axis_title = "Item Pertanyaan",
                                    y_axis_title = "",
                                    y_metric = "Percentage",
                                    label_mode = "Percentage",
                                    orientation = "Vertical",
                                    theme_name = "White Publication",
                                    bar_width = 0.72,
                                    show_labels = TRUE,
                                    label_size = 3.2,
                                    label_color = "#111111",
                                    label_vjust = 0.5,
                                    facet_by_split = TRUE,
                                    split_panel_cols = 1,
                                    facet_scales = "fixed",
                                    legend_position = "Right",
                                    title_size = 16,
                                    subtitle_size = 11,
                                    axis_title_size = 11,
                                    axis_text_size = 9,
                                    legend_title_size = 10,
                                    legend_text_size = 9,
                                    x_text_angle = 35,
                                    digits = 2) {
  validate(need(nrow(chart_df) > 0, "The selected Likert data are empty or do not match the selected response order."))
  th <- get_theme(theme_name)
  chart_df$Y_Value <- if (y_metric == "Frequency") chart_df$Frequency else chart_df$Percentage
  chart_df$Response <- factor(as.character(chart_df$Response), levels = levels)
  chart_df$Label <- ""
  
  if (label_mode == "Frequency") {
    chart_df$Label <- ifelse(chart_df$Frequency > 0, as.character(chart_df$Frequency), "")
  } else if (label_mode == "Percentage") {
    chart_df$Label <- ifelse(chart_df$Frequency > 0, paste0(format(round(chart_df$Percentage, digits), nsmall = digits), "%"), "")
  } else if (label_mode == "Frequency and Percentage") {
    chart_df$Label <- ifelse(
      chart_df$Frequency > 0,
      paste0(chart_df$Frequency, " (", format(round(chart_df$Percentage, digits), nsmall = digits), "%)"),
      ""
    )
  }
  
  default_y_label <- if (y_metric == "Frequency") "Frequency" else "Percentage (%)"
  y_label <- if (!is.null(y_axis_title) && nzchar(trimws(y_axis_title))) y_axis_title else default_y_label
  x_label <- if (!is.null(x_axis_title) && nzchar(trimws(x_axis_title))) x_axis_title else "Item Pertanyaan"
  legend_title <- if (is.null(legend_title)) "" else as.character(legend_title)
  
  if (is.null(legend_labels) || length(legend_labels) != length(levels)) {
    legend_labels <- levels
    names(legend_labels) <- levels
  }
  
  p <- ggplot(chart_df, aes(x = `Item Pertanyaan`, y = Y_Value, fill = Response)) +
    geom_col(width = bar_width, color = "white", linewidth = 0.25) +
    scale_fill_manual(
      values = manual_colors,
      breaks = levels,
      labels = unname(legend_labels[levels]),
      drop = FALSE
    ) +
    labs(title = title, subtitle = subtitle, x = x_label, y = y_label, fill = legend_title) +
    statcal_theme_gg(
      theme_name = theme_name,
      title_size = title_size,
      subtitle_size = subtitle_size,
      axis_title_size = axis_title_size,
      axis_text_size = axis_text_size,
      legend_title_size = legend_title_size,
      legend_text_size = legend_text_size,
      legend_position = legend_position,
      x_text_angle = x_text_angle
    )
  
  if (show_labels && label_mode != "None") {
    p <- p + geom_text(
      aes(label = Label),
      position = position_stack(vjust = label_vjust),
      size = label_size,
      color = label_color,
      check_overlap = TRUE
    )
  }
  
  if (facet_by_split && any(as.character(chart_df$Split) != "All")) {
    facet_scales <- ifelse(is.null(facet_scales) || !(facet_scales %in% c("fixed", "free_y", "free_x", "free")), "fixed", facet_scales)
    p <- p + facet_wrap(~ Split, ncol = split_panel_cols, scales = facet_scales)
  }
  
  if (orientation == "Horizontal") p <- p + coord_flip()
  
  p + theme(plot.background = element_rect(fill = th$figure_facecolor, color = NA))
}


# ============================================================
# AUTOMATIC WORD INTERPRETATION HELPERS
# ============================================================
# These functions create descriptive narratives from the same long-format
# Likert distribution objects used by the tables and stacked bar chart.
# The narrative is descriptive only and does not imply statistical
# significance, causality, or inferential differences between groups.

safe_pct_text <- function(x, digits = 2) {
  x <- suppressWarnings(as.numeric(x))
  if (length(x) == 0 || is.na(x) || !is.finite(x)) return("NA")
  paste0(format(round(x, digits), nsmall = digits, trim = TRUE, scientific = FALSE), "%")
}

response_display_label <- function(response_value, legend_labels) {
  response_value <- as.character(response_value)
  if (!is.null(legend_labels) && response_value %in% names(legend_labels)) {
    out <- as.character(legend_labels[[response_value]])
    if (length(out) > 0 && !is.na(out) && nzchar(trimws(out))) return(trimws(out))
  }
  response_value
}

add_word_paragraphs <- function(doc, paragraphs) {
  if (is.null(paragraphs) || length(paragraphs) == 0) return(doc)
  for (txt in paragraphs) {
    if (!is.null(txt) && length(txt) > 0 && !is.na(txt) && nzchar(trimws(as.character(txt)))) {
      doc <- officer::body_add_par(doc, as.character(txt), style = "Normal")
    }
  }
  doc
}

build_likert_word_narratives <- function(overall_long_df,
                                         split_long_df,
                                         levels,
                                         legend_labels,
                                         split_var = "None",
                                         digits = 2,
                                         chart_metric = "Percentage",
                                         facet_by_split = TRUE,
                                         language = "Bahasa Indonesia") {
  
  overall_long_df <- as.data.frame(overall_long_df)
  split_long_df <- as.data.frame(split_long_df)
  levels <- as.character(levels)
  
  if (length(levels) == 0 || nrow(overall_long_df) == 0) {
    return(list(
      table1 = character(0),
      table2 = character(0),
      figure1 = character(0),
      table3 = character(0)
    ))
  }
  
  is_id <- identical(language, "Bahasa Indonesia")
  
  # ---------- Overall distribution ----------
  overall_resp <- overall_long_df %>%
    dplyr::group_by(Response) %>%
    dplyr::summarise(Frequency = sum(Frequency, na.rm = TRUE), .groups = "drop")
  
  total_item_responses <- sum(overall_resp$Frequency, na.rm = TRUE)
  overall_resp$Share <- if (total_item_responses > 0) {
    overall_resp$Frequency / total_item_responses * 100
  } else {
    NA_real_
  }
  
  overall_resp$Response_chr <- as.character(overall_resp$Response)
  overall_resp$Response_chr <- factor(overall_resp$Response_chr, levels = levels)
  overall_resp <- overall_resp[order(overall_resp$Response_chr), , drop = FALSE]
  
  dom_overall <- overall_resp[order(-overall_resp$Share), , drop = FALSE][1, , drop = FALSE]
  dom_overall_label <- response_display_label(dom_overall$Response[1], legend_labels)
  dom_overall_pct <- safe_pct_text(dom_overall$Share[1], digits)
  
  item_order <- base::levels(overall_long_df$`Item Pertanyaan`)
  if (is.null(item_order) || length(item_order) == 0) {
    item_order <- unique(as.character(overall_long_df$`Item Pertanyaan`))
  }
  
  item_dom <- overall_long_df %>%
    dplyr::mutate(
      Item_chr = as.character(`Item Pertanyaan`),
      Response_chr = as.character(Response)
    ) %>%
    dplyr::group_by(Item_chr) %>%
    dplyr::arrange(dplyr::desc(Percentage), .by_group = TRUE) %>%
    dplyr::slice_head(n = 1) %>%
    dplyr::ungroup()
  
  dom_counts <- item_dom %>%
    dplyr::count(Response_chr, name = "N_Items") %>%
    dplyr::arrange(dplyr::desc(N_Items), Response_chr)
  
  most_common_dom <- if (nrow(dom_counts) > 0) dom_counts[1, , drop = FALSE] else NULL
  strongest_item <- if (nrow(item_dom) > 0) item_dom[order(-item_dom$Percentage), , drop = FALSE][1, , drop = FALSE] else NULL
  
  # ---------- Table 1 narrative ----------
  table1 <- character(0)
  
  if (is_id) {
    table1 <- c(
      table1,
      paste0(
        "Tabel distribusi Likert merangkum ", length(item_order),
        " item pertanyaan dengan total ", format(total_item_responses, big.mark = ".", scientific = FALSE),
        " respons item yang valid. Secara keseluruhan, kategori respons yang paling banyak muncul adalah \"",
        dom_overall_label, "\" dengan proporsi ", dom_overall_pct, "."
      )
    )
    
    if (!is.null(most_common_dom) && nrow(most_common_dom) > 0) {
      lbl <- response_display_label(most_common_dom$Response_chr[1], legend_labels)
      table1 <- c(
        table1,
        paste0(
          "Jika dilihat berdasarkan kategori dengan persentase terbesar pada masing-masing item, respons \"",
          lbl, "\" paling sering menjadi kategori dominan, yaitu pada ",
          most_common_dom$N_Items[1], " dari ", length(item_order), " item."
        )
      )
    }
    
    if (!is.null(strongest_item) && nrow(strongest_item) > 0) {
      lbl <- response_display_label(strongest_item$Response_chr[1], legend_labels)
      table1 <- c(
        table1,
        paste0(
          "Dominasi kategori respons paling kuat terdapat pada item \"",
          strongest_item$Item_chr[1], "\", dengan kategori \"", lbl,
          "\" sebesar ", safe_pct_text(strongest_item$Percentage[1], digits),
          ". Hasil ini bersifat deskriptif dan menunjukkan pola distribusi jawaban pada data yang dianalisis."
        )
      )
    }
  } else {
    table1 <- c(
      table1,
      paste0(
        "The Likert distribution table summarizes ", length(item_order),
        " question items with a total of ", format(total_item_responses, big.mark = ",", scientific = FALSE),
        " valid item responses. Overall, the most frequent response category is \"",
        dom_overall_label, "\" with a share of ", dom_overall_pct, "."
      )
    )
    
    if (!is.null(most_common_dom) && nrow(most_common_dom) > 0) {
      lbl <- response_display_label(most_common_dom$Response_chr[1], legend_labels)
      table1 <- c(
        table1,
        paste0(
          "Based on the largest percentage within each item, the response \"",
          lbl, "\" is the most frequently dominant category, leading in ",
          most_common_dom$N_Items[1], " of ", length(item_order), " items."
        )
      )
    }
    
    if (!is.null(strongest_item) && nrow(strongest_item) > 0) {
      lbl <- response_display_label(strongest_item$Response_chr[1], legend_labels)
      table1 <- c(
        table1,
        paste0(
          "The strongest single-item concentration occurs for item \"",
          strongest_item$Item_chr[1], "\", where the category \"", lbl,
          "\" accounts for ", safe_pct_text(strongest_item$Percentage[1], digits),
          ". These results are descriptive and summarize the observed response distribution."
        )
      )
    }
  }
  
  # ---------- Table 2 narrative ----------
  use_split <- !is.null(split_var) && split_var != "None" &&
    nrow(split_long_df) > 0 &&
    any(as.character(split_long_df$Split) != "All")
  
  table2 <- character(0)
  
  if (!use_split) {
    if (is_id) {
      table2 <- c(
        "Variabel pemisah belum dipilih. Oleh karena itu, tabel berdasarkan split menampilkan distribusi yang sama dengan tabel frekuensi keseluruhan dan belum memberikan perbandingan antar-kelompok."
      )
    } else {
      table2 <- c(
        "No split variable is selected. Therefore, the split table represents the same overall distribution as the frequency table and does not yet provide a between-group comparison."
      )
    }
  } else {
    split_summary <- split_long_df %>%
      dplyr::mutate(
        Split_chr = as.character(Split),
        Response_chr = as.character(Response)
      ) %>%
      dplyr::group_by(Split_chr, Response_chr) %>%
      dplyr::summarise(Frequency = sum(Frequency, na.rm = TRUE), .groups = "drop") %>%
      dplyr::group_by(Split_chr) %>%
      dplyr::mutate(
        GroupTotal = sum(Frequency, na.rm = TRUE),
        Share = ifelse(GroupTotal > 0, Frequency / GroupTotal * 100, NA_real_)
      ) %>%
      dplyr::ungroup()
    
    split_dom <- split_summary %>%
      dplyr::group_by(Split_chr) %>%
      dplyr::arrange(dplyr::desc(Share), .by_group = TRUE) %>%
      dplyr::slice_head(n = 1) %>%
      dplyr::ungroup()
    
    split_levels <- unique(as.character(split_long_df$Split))
    split_levels <- split_levels[!is.na(split_levels) & nzchar(split_levels)]
    
    max_groups_to_narrate <- min(length(split_levels), 8L)
    group_sentences <- character(0)
    
    if (nrow(split_dom) > 0) {
      for (i in seq_len(min(nrow(split_dom), max_groups_to_narrate))) {
        lbl <- response_display_label(split_dom$Response_chr[i], legend_labels)
        if (is_id) {
          group_sentences <- c(
            group_sentences,
            paste0(
              split_dom$Split_chr[i], ": kategori dominan \"", lbl,
              "\" (", safe_pct_text(split_dom$Share[i], digits), ")"
            )
          )
        } else {
          group_sentences <- c(
            group_sentences,
            paste0(
              split_dom$Split_chr[i], ": dominant category \"", lbl,
              "\" (", safe_pct_text(split_dom$Share[i], digits), ")"
            )
          )
        }
      }
    }
    
    last_level <- tail(levels, 1)
    last_label <- response_display_label(last_level, legend_labels)
    
    last_level_share <- split_summary[
      split_summary$Response_chr == last_level,
      c("Split_chr", "Share"),
      drop = FALSE
    ]
    
    if (is_id) {
      table2 <- c(
        table2,
        paste0(
          "Tabel split membandingkan distribusi jawaban berdasarkan variabel \"",
          split_var, "\" yang terdiri atas ", length(split_levels), " kelompok. ",
          "Kategori respons dominan pada masing-masing kelompok adalah: ",
          paste(group_sentences, collapse = "; "), "."
        )
      )
    } else {
      table2 <- c(
        table2,
        paste0(
          "The split table compares response distributions across the variable \"",
          split_var, "\", consisting of ", length(split_levels), " groups. ",
          "The dominant response category in each group is: ",
          paste(group_sentences, collapse = "; "), "."
        )
      )
    }
    
    if (nrow(last_level_share) > 1 && any(is.finite(last_level_share$Share))) {
      max_row <- last_level_share[which.max(last_level_share$Share), , drop = FALSE]
      min_row <- last_level_share[which.min(last_level_share$Share), , drop = FALSE]
      
      if (is_id) {
        table2 <- c(
          table2,
          paste0(
            "Untuk kategori respons terakhir dalam urutan skala, yaitu \"", last_label,
            "\", proporsi terbesar terdapat pada kelompok \"", max_row$Split_chr[1],
            "\" sebesar ", safe_pct_text(max_row$Share[1], digits),
            ", sedangkan proporsi terkecil terdapat pada kelompok \"",
            min_row$Split_chr[1], "\" sebesar ", safe_pct_text(min_row$Share[1], digits),
            ". Perbedaan ini merupakan perbandingan deskriptif dan tidak menunjukkan signifikansi statistik."
          )
        )
      } else {
        table2 <- c(
          table2,
          paste0(
            "For the last response category in the specified scale order, \"", last_label,
            "\", the largest share is observed in group \"", max_row$Split_chr[1],
            "\" at ", safe_pct_text(max_row$Share[1], digits),
            ", while the smallest share is observed in group \"",
            min_row$Split_chr[1], "\" at ", safe_pct_text(min_row$Share[1], digits),
            ". This is a descriptive comparison and does not imply statistical significance."
          )
        )
      }
    }
  }
  
  # ---------- Figure narrative ----------
  metric_text_id <- if (identical(chart_metric, "Frequency")) "frekuensi" else "persentase"
  metric_text_en <- if (identical(chart_metric, "Frequency")) "frequency" else "percentage"
  
  if (is_id) {
    figure1 <- c(
      paste0(
        "Gambar stacked bar memvisualisasikan distribusi respons Likert berdasarkan ",
        metric_text_id, ". Setiap warna merepresentasikan satu kategori respons sesuai urutan skala yang telah ditentukan. ",
        "Segmen yang lebih besar menunjukkan frekuensi atau proporsi respons yang lebih besar pada item yang bersangkutan."
      )
    )
    if (use_split && isTRUE(facet_by_split)) {
      figure1 <- c(
        figure1,
        paste0(
          "Grafik ditampilkan dalam panel berdasarkan variabel \"", split_var,
          "\", sehingga pola distribusi antar-kelompok dapat dibandingkan secara visual menggunakan sumber data yang sama dengan tabel split."
        )
      )
    }
    figure1 <- c(
      figure1,
      paste0(
        "Secara keseluruhan, kategori \"", dom_overall_label,
        "\" merupakan respons yang paling dominan pada data terpilih dengan proporsi ",
        dom_overall_pct, ". Interpretasi grafik ini bersifat deskriptif."
      )
    )
  } else {
    figure1 <- c(
      paste0(
        "The stacked bar chart visualizes the Likert response distribution using ",
        metric_text_en, ". Each color represents one response category according to the specified scale order. ",
        "Larger segments indicate a greater frequency or proportion of responses for the corresponding item."
      )
    )
    if (use_split && isTRUE(facet_by_split)) {
      figure1 <- c(
        figure1,
        paste0(
          "The chart is displayed in panels based on the variable \"", split_var,
          "\", allowing visual comparison of response patterns across groups using the same distribution data as the split table."
        )
      )
    }
    figure1 <- c(
      figure1,
      paste0(
        "Overall, the category \"", dom_overall_label,
        "\" is the most dominant response in the selected data, accounting for ",
        dom_overall_pct, ". The chart interpretation is descriptive."
      )
    )
  }
  
  # ---------- Table 3 narrative ----------
  if (is_id) {
    table3 <- c(
      paste0(
        "Tabel data grafik disajikan dalam format panjang dan menjadi sumber langsung pembentukan stacked bar chart. ",
        "Setiap baris merepresentasikan kombinasi item pertanyaan, kelompok split, dan kategori respons, disertai nilai frekuensi, total respons valid, serta persentase. ",
        "Dengan demikian, angka pada tabel ini konsisten dengan nilai yang digunakan dalam visualisasi."
      )
    )
  } else {
    table3 <- c(
      paste0(
        "The chart-data table is presented in long format and serves as the direct source for the stacked bar chart. ",
        "Each row represents a combination of question item, split group, and response category, together with frequency, valid-response total, and percentage. ",
        "Therefore, the values in this table are synchronized with those used in the visualization."
      )
    )
  }
  
  list(
    table1 = table1,
    table2 = table2,
    figure1 = figure1,
    table3 = table3
  )
}


# ============================================================
# ROBUST EXPORT HELPERS - SESSION TEMP FILE + BROWSER BLOB
# ============================================================

make_export_filename <- function(prefix, ext = "png", dpi = NULL) {
  stamp <- format(Sys.time(), "%Y%m%d_%H%M%S")
  if (!is.null(dpi)) paste0(prefix, "_", dpi, "dpi_", stamp, ".", ext) else paste0(prefix, "_", stamp, ".", ext)
}

create_session_export_dir <- function(session_token) {
  token <- ifelse(is.null(session_token) || !nzchar(session_token), paste0(Sys.getpid()), safe_token(session_token))
  out <- file.path(tempdir(), paste0("statcal_likert_", token))
  if (!dir.exists(out)) dir.create(out, recursive = TRUE, showWarnings = FALSE)
  out
}

validate_export_file <- function(path, label = "Export file") {
  if (!file.exists(path)) stop(label, " was not created.")
  size <- file.info(path)$size
  if (is.na(size) || size <= 0) stop(label, " is empty.")
  invisible(TRUE)
}

generate_plot_png <- function(plot_object, file, width = 8, height = 6, dpi = 1200, bg = "white") {
  width <- safe_number(width, 8, 3, 30)
  height <- safe_number(height, 6, 3, 30)
  dpi <- safe_number(dpi, 1200, 72, 1500)
  if (!inherits(plot_object, "ggplot")) stop("The selected chart is not a ggplot object.")
  ggplot2::ggsave(
    filename = file,
    plot = plot_object,
    width = width,
    height = height,
    units = "in",
    dpi = dpi,
    bg = bg,
    limitsize = FALSE
  )
  validate_export_file(file, "PNG file")
  invisible(file)
}

new_export_result <- function(path, filename, message) {
  validate_export_file(path)
  list(
    ok = TRUE,
    file = normalizePath(path, winslash = "/", mustWork = FALSE),
    filename = filename,
    message = message,
    size = file.info(path)$size
  )
}

# Send an already-generated binary file directly to the browser as a Blob.
# This bypasses downloadHandler URLs and the Shinylive Service Worker
# download routing that can produce "Couldn't find parent page" errors.
send_binary_file_to_browser <- function(session,
                                        path,
                                        filename,
                                        mime = "application/octet-stream",
                                        max_mb = 25) {
  validate_export_file(path)
  
  size_bytes <- file.info(path)$size
  if (is.na(size_bytes) || size_bytes <= 0) {
    stop("The generated file is empty.")
  }
  
  if (size_bytes > max_mb * 1024^2) {
    stop(
      sprintf(
        paste0(
          "The generated file is %.1f MB, which is too large for the ",
          "browser Blob transfer (limit %.0f MB). Reduce PNG DPI/size ",
          "or omit very large raw-data appendices and try again."
        ),
        size_bytes / 1024^2,
        max_mb
      )
    )
  }
  
  encoded <- base64enc::base64encode(path, linewidth = 0)
  
  session$sendCustomMessage(
    "statcal_binary_download",
    list(
      filename = filename,
      mime = mime,
      data = encoded
    )
  )
  
  invisible(TRUE)
}

generated_download_ui <- function(result, download_id, button_label, icon_name = "download") {
  if (is.null(result)) {
    return(
      tags$p(
        class = "small-note",
        "Click Generate first. After the file has been created and verified, the download button will appear."
      )
    )
  }
  
  if (!isTRUE(result$ok)) {
    return(tags$div(class = "alert alert-danger", result$message))
  }
  
  tagList(
    tags$div(
      class = "alert alert-success",
      tags$b(result$message),
      tags$br(),
      tags$span(sprintf("File: %s (%.1f KB)", result$filename, result$size / 1024))
    ),
    actionButton(
      download_id,
      tagList(icon(icon_name), button_label),
      class = "btn-primary"
    )
  )
}

safe_sheet_name <- function(name) {
  name <- as.character(name)
  invalid_chars <- c("\\", "/", "?", "*", "[", "]", ":")
  for (ch in invalid_chars) name <- gsub(ch, "_", name, fixed = TRUE)
  name <- trimws(name)
  name <- substr(name, 1, 31)
  ifelse(nchar(name) == 0, "Sheet", name)
}

write_table_sheet <- function(wb, sheet_name, df, table_title = NULL) {
  sheet_name <- safe_sheet_name(sheet_name)
  openxlsx::addWorksheet(wb, sheet_name)
  df <- make_display_safe(as.data.frame(df))
  start_row <- 1
  if (!is.null(table_title) && nzchar(trimws(table_title))) {
    openxlsx::writeData(wb, sheet_name, table_title, startRow = 1, startCol = 1)
    title_style <- openxlsx::createStyle(textDecoration = "bold", fontSize = 13, fontColour = "#1F4E79")
    openxlsx::addStyle(wb, sheet_name, title_style, rows = 1, cols = 1, stack = TRUE)
    start_row <- 3
  }
  openxlsx::writeData(wb, sheet_name, df, startRow = start_row)
  if (ncol(df) > 0) {
    openxlsx::setColWidths(wb, sheet_name, cols = 1:ncol(df), widths = "auto")
    header_style <- openxlsx::createStyle(textDecoration = "bold", fgFill = "#D9EAF7", border = "Bottom")
    openxlsx::addStyle(wb, sheet_name, header_style, rows = start_row, cols = 1:ncol(df), gridExpand = TRUE)
    openxlsx::freezePane(wb, sheet_name, firstActiveRow = start_row + 1)
  }
}

export_likert_workbook <- function(file, raw_df, frequency_df, split_df, chart_df, metadata_df,
                                   plot_object, chart_title,
                                   table1_title, table2_title, table3_title,
                                   export_width = 8, export_height = 6) {
  wb <- openxlsx::createWorkbook()
  write_table_sheet(wb, "Export Info", metadata_df, "STATCAL ONLINE - Export Information")
  write_table_sheet(wb, "Raw Data", raw_df, "Appendix A. Raw Data")
  write_table_sheet(wb, "Frequency Table", frequency_df, table1_title)
  write_table_sheet(wb, "Split Frequency Table", split_df, table2_title)
  write_table_sheet(wb, "Chart Data", chart_df, table3_title)
  
  tmp_png <- tempfile(fileext = ".png")
  on.exit(unlink(tmp_png), add = TRUE)
  generate_plot_png(plot_object, tmp_png, width = export_width, height = export_height, dpi = 300, bg = "white")
  openxlsx::addWorksheet(wb, "Figure 1")
  openxlsx::writeData(wb, "Figure 1", chart_title, startRow = 1, startCol = 1)
  fig_style <- openxlsx::createStyle(textDecoration = "bold", fontSize = 13, fontColour = "#1F4E79")
  openxlsx::addStyle(wb, "Figure 1", fig_style, rows = 1, cols = 1)
  openxlsx::insertImage(wb, "Figure 1", tmp_png, startRow = 3, startCol = 1,
                        width = min(safe_number(export_width, 8, 4, 16), 12),
                        height = min(safe_number(export_height, 6, 3, 14), 10), units = "in")
  
  openxlsx::saveWorkbook(wb, file, overwrite = TRUE)
  validate_export_file(file, "Excel workbook")
}

make_report_flextable <- function(df, font_size = 8) {
  df <- make_display_safe(as.data.frame(df))
  ft <- flextable::flextable(df)
  ft <- flextable::theme_booktabs(ft)
  ft <- flextable::bold(ft, part = "header")
  ft <- flextable::bg(ft, bg = "#D9EAF7", part = "header")
  ft <- flextable::fontsize(ft, size = font_size, part = "all")
  ft <- flextable::autofit(ft)
  ft
}

export_word_report <- function(file, report_title, report_subtitle,
                               metadata_df, frequency_df, split_df, chart_df, raw_df,
                               plot_object, table1_title, table2_title, table3_title, figure1_title,
                               include_chart_data = TRUE, include_raw_data = FALSE,
                               include_interpretation = TRUE,
                               table1_narrative = character(0),
                               table2_narrative = character(0),
                               figure1_narrative = character(0),
                               table3_narrative = character(0),
                               plot_width = 6.4, plot_height = 4.5) {
  doc <- officer::read_docx()
  doc <- officer::body_add_par(doc, report_title, style = "heading 1")
  if (!is.null(report_subtitle) && nzchar(trimws(report_subtitle))) {
    doc <- officer::body_add_par(doc, report_subtitle, style = "Normal")
  }
  doc <- officer::body_add_par(doc, paste("Generated by STATCAL ONLINE on", format(Sys.time(), "%d %B %Y %H:%M")), style = "Normal")
  
  doc <- officer::body_add_par(doc, "Analysis Settings", style = "heading 2")
  doc <- flextable::body_add_flextable(doc, make_report_flextable(metadata_df, 8))
  
  doc <- officer::body_add_par(doc, table1_title, style = "heading 2")
  doc <- flextable::body_add_flextable(doc, make_report_flextable(frequency_df, 7.5))
  if (isTRUE(include_interpretation)) {
    doc <- add_word_paragraphs(doc, table1_narrative)
  }
  
  doc <- officer::body_add_par(doc, table2_title, style = "heading 2")
  doc <- flextable::body_add_flextable(doc, make_report_flextable(split_df, 7.2))
  if (isTRUE(include_interpretation)) {
    doc <- add_word_paragraphs(doc, table2_narrative)
  }
  
  tmp_png <- tempfile(fileext = ".png")
  on.exit(unlink(tmp_png), add = TRUE)
  generate_plot_png(plot_object, tmp_png, width = plot_width, height = plot_height, dpi = 300, bg = "white")
  doc <- officer::body_add_par(doc, figure1_title, style = "heading 2")
  doc <- officer::body_add_img(doc, src = tmp_png, width = plot_width, height = plot_height)
  if (isTRUE(include_interpretation)) {
    doc <- add_word_paragraphs(doc, figure1_narrative)
  }
  
  if (isTRUE(include_chart_data)) {
    doc <- officer::body_add_par(doc, table3_title, style = "heading 2")
    doc <- flextable::body_add_flextable(doc, make_report_flextable(chart_df, 7.2))
    if (isTRUE(include_interpretation)) {
      doc <- add_word_paragraphs(doc, table3_narrative)
    }
  }
  
  if (isTRUE(include_raw_data)) {
    doc <- officer::body_add_break(doc)
    doc <- officer::body_add_par(doc, "Appendix A. Raw Data", style = "heading 2")
    doc <- flextable::body_add_flextable(doc, make_report_flextable(raw_df, 6.8))
  }
  
  print(doc, target = file)
  validate_export_file(file, "Word report")
}

# ============================================================
# UI
# ============================================================

legend_choices <- c("Right", "Left", "Top", "Bottom", "None / Hide legend")

ui <- dashboardPage(
  dashboardHeader(title = APP_NAME, titleWidth = "100%"),
  dashboardSidebar(disable = TRUE),
  dashboardBody(
    tags$head(
      tags$style(HTML("\n        .content-wrapper, .right-side { background-color: #f7f9fb; }\n        .box { border-radius: 10px; }\n        .statcal-title { font-size: 24px; font-weight: 700; color: #1F4E79; }\n        .statcal-subtitle { font-size: 18px; font-weight: 600; color: #333333; }\n        .statcal-note { line-height: 1.6; text-align: justify; }\n        .small-note { font-size: 12px; color: #666666; }\n      ")),
      tags$script(HTML("
        (function registerStatcalBinaryDownload() {
          if (!window.Shiny || !Shiny.addCustomMessageHandler) {
            window.setTimeout(registerStatcalBinaryDownload, 50);
            return;
          }

          Shiny.addCustomMessageHandler('statcal_binary_download', function(message) {
            try {
              var binary = window.atob(message.data || '');
              var len = binary.length;
              var bytes = new Uint8Array(len);

              for (var i = 0; i < len; i++) {
                bytes[i] = binary.charCodeAt(i);
              }

              var blob = new Blob(
                [bytes],
                { type: message.mime || 'application/octet-stream' }
              );

              var url = window.URL.createObjectURL(blob);
              var a = document.createElement('a');

              a.style.display = 'none';
              a.href = url;
              a.download = message.filename || 'statcal_download.bin';

              document.body.appendChild(a);
              a.click();

              window.setTimeout(function() {
                window.URL.revokeObjectURL(url);
                if (a.parentNode) {
                  a.parentNode.removeChild(a);
                }
              }, 1500);

            } catch (err) {
              console.error('STATCAL binary download failed:', err);
              window.alert('Download failed in the browser: ' + err.message);
            }
          });
        })();
      "))
    ),
    fluidRow(
      box(
        width = 12, status = "primary", solidHeader = TRUE,
        title = "STATCAL ONLINE for Likert Scale Distribution and Stacked Bar Chart Visualizer",
        div(class = "statcal-title", APP_TITLE),
        div(class = "statcal-subtitle", APP_UPDATED),
        tags$p(class = "statcal-note",
               "This R Shiny application is designed to analyze Likert-scale survey items. Users can select question items flexibly, define the order of Likert labels, generate frequency and percentage distribution tables, split the analysis by a categorical variable, create publication-ready stacked bar charts, and export tables or charts for academic reporting."
        ),
        tags$p(
          tags$b("Website: "), tags$a(href = WEBSITE_URL, target = "_blank", WEBSITE_URL), tags$br(),
          tags$b("STATCAL ONLINE Page: "), tags$a(href = STATCAL_ONLINE_URL, target = "_blank", STATCAL_ONLINE_URL), tags$br(),
          tags$b("Training Data: "), tags$a(href = URL_DATA_TRAINING, target = "_blank", "Open Google Drive Folder")
        )
      )
    ),
    tabsetPanel(
      id = "main_tabs",
      tabPanel(
        "1. Data & Settings",
        br(),
        fluidRow(
          box(width = 5, title = "Data Input", status = "primary", solidHeader = TRUE,
              fileInput("uploaded_file", "Upload Excel file", accept = c(".xlsx", ".xls")),
              uiOutput("sheet_ui"),
              tags$p(class = "small-note", "If no file is uploaded, the application uses the sample data_likert.xlsx file.")),
          box(width = 7, title = "Likert Item and Response Order Settings", status = "primary", solidHeader = TRUE,
              uiOutput("item_vars_ui"),
              textInput("likert_order_text", "Likert response order (comma-separated)", value = DEFAULT_LIKERT_ORDER),
              sliderInput("decimal_digits", "Decimal digits for percentages", min = 0, max = 8, value = 2, step = 1),
              uiOutput("split_var_ui"),
              uiOutput("split_order_ui"))
        ),
        fluidRow(
          valueBoxOutput("metric_rows", width = 3),
          valueBoxOutput("metric_columns", width = 3),
          valueBoxOutput("metric_items", width = 3),
          valueBoxOutput("metric_levels", width = 3)
        ),
        fluidRow(
          box(width = 12, title = "Dataset Preview", status = "warning", solidHeader = TRUE,
              shinycssloaders::withSpinner(DTOutput("data_preview")))
        )
      ),
      tabPanel(
        "2. Frequency Table",
        br(),
        fluidRow(
          box(width = 12, title = "Likert Frequency and Percentage Table", status = "warning", solidHeader = TRUE,
              tags$p("Table format: f = frequency, % = percentage. Percentages are calculated within each selected item question."),
              shinycssloaders::withSpinner(DTOutput("frequency_table")))
        )
      ),
      tabPanel(
        "3. Split Table",
        br(),
        fluidRow(
          box(width = 12, title = "Likert Frequency and Percentage Table by Split Variable", status = "warning", solidHeader = TRUE,
              tags$p("Select a split variable in Data & Settings, for example Jenis Kelamin or Pendidikan. This table and the stacked bar chart are generated from the same distribution dataset, so frequency and percentage values are synchronized."),
              shinycssloaders::withSpinner(DTOutput("split_frequency_table")))
        )
      ),
      tabPanel(
        "4. Stacked Bar Chart",
        br(),
        fluidRow(
          box(width = 3, title = "Chart Data", status = "primary", solidHeader = TRUE,
              selectInput("chart_y_metric", "Bar height", choices = c("Frequency", "Percentage"), selected = "Percentage"),
              selectInput("chart_label_mode", "Show information on bars", choices = c("None", "Frequency", "Percentage", "Frequency and Percentage"), selected = "Percentage"),
              selectInput("chart_orientation", "Bar orientation", choices = c("Vertical", "Horizontal"), selected = "Vertical"),
              checkboxInput("chart_facet_split", "Create panels by split variable", value = TRUE),
              sliderInput("chart_split_cols", "Split panel columns (1 = vertical panels)", min = 1, max = 4, value = 1, step = 1),
              selectInput("chart_facet_scales", "Panel axis scale",
                          choices = c("Same scale across panels" = "fixed", "Free Y scale by panel" = "free_y", "Free X scale by panel" = "free_x", "Free X and Y scale" = "free"),
                          selected = "fixed")),
          box(width = 3, title = "Bar & Palette Style", status = "primary", solidHeader = TRUE,
              selectInput("chart_theme", "Background theme", choices = names(THEMES), selected = "Minimal Scopus Style"),
              selectInput("chart_palette", "Bar color palette", choices = names(LIKERT_PALETTES), selected = "Scopus-inspired Orange–Blue"),
              checkboxInput("chart_reverse_palette", "Reverse selected palette", value = FALSE),
              uiOutput("palette_preview_ui"),
              sliderInput("bar_width", "Bar width", min = 0.25, max = 1.00, value = 0.72, step = 0.01),
              sliderInput("chart_height_px", "Preview chart height (px)", min = 350, max = 1200, value = 700, step = 50)),
          box(width = 3, title = "Titles & Data Labels", status = "primary", solidHeader = TRUE,
              textInput("chart_title", "Chart title", value = "Likert Scale Distribution"),
              textInput("chart_subtitle", "Chart subtitle", value = "Stacked bar chart based on selected survey items"),
              textInput("chart_x_axis_title", "X-axis title", value = "Item Pertanyaan"),
              textInput("chart_y_axis_title", "Y-axis title (blank = automatic)", value = ""),
              checkboxInput("chart_show_labels", "Show text labels", value = TRUE),
              sliderInput("chart_label_size", "Label text size", min = 2, max = 8, value = 3.2, step = 0.2),
              colourpicker::colourInput("chart_label_color", "Label text color", value = "#111111", showColour = "both"),
              sliderInput("chart_label_vjust", "Label position inside stack", min = 0.05, max = 0.95, value = 0.50, step = 0.05)),
          box(width = 3, title = "Legend Settings", status = "primary", solidHeader = TRUE,
              selectInput("chart_legend_position", "Legend position", choices = legend_choices, selected = "Right"),
              textInput("chart_legend_title", "Legend title", value = "Likert Response"),
              tags$p(class = "small-note", "Change the legend display labels without changing the original data values."),
              uiOutput("legend_label_settings_ui"),
              sliderInput("chart_x_text_angle", "X-axis text angle", min = 0, max = 90, value = 35, step = 5))
        ),
        fluidRow(
          box(width = 6, title = "Manual / Custom Bar Colors", status = "info", solidHeader = TRUE, collapsible = TRUE, collapsed = TRUE,
              tags$p(class = "small-note", "Select 'Manual / Custom' as the palette to activate these color pickers."),
              uiOutput("color_settings_ui")),
          box(width = 6, title = "Flexible Text Size Settings", status = "info", solidHeader = TRUE, collapsible = TRUE, collapsed = TRUE,
              fluidRow(
                column(4, sliderInput("chart_title_size", "Title", 8, 34, 16, 1)),
                column(4, sliderInput("chart_subtitle_size", "Subtitle", 6, 26, 11, 1)),
                column(4, sliderInput("chart_axis_title_size", "Axis title", 6, 24, 11, 1)),
                column(4, sliderInput("chart_axis_text_size", "Axis text", 5, 22, 9, 1)),
                column(4, sliderInput("chart_legend_title_size", "Legend title", 5, 24, 10, 1)),
                column(4, sliderInput("chart_legend_text_size", "Legend text", 5, 22, 9, 1))
              ))
        ),
        fluidRow(
          box(width = 12, title = "Publication-Ready Stacked Bar Chart", status = "warning", solidHeader = TRUE,
              shinycssloaders::withSpinner(uiOutput("stacked_plot_ui")))
        ),
        fluidRow(
          box(width = 12, title = "Chart Data", status = "info", solidHeader = TRUE, collapsible = TRUE, collapsed = TRUE,
              shinycssloaders::withSpinner(DTOutput("chart_data_table")))
        )
      ),
      tabPanel(
        "5. Export Excel, Word & Figure",
        br(),
        fluidRow(
          box(width = 12, title = "Export Settings and Captions", status = "primary", solidHeader = TRUE,
              fluidRow(
                column(4, selectInput("export_dpi", "PNG resolution / DPI", choices = c(300, 600, 900, 1200, 1500), selected = 1200)),
                column(4, numericInput("export_width", "Export width (inches)", value = 8, min = 4, max = 30, step = 0.5)),
                column(4, numericInput("export_height", "Export height (inches)", value = 6, min = 3, max = 30, step = 0.5))
              ),
              fluidRow(
                column(6, textInput("word_report_title", "Word report title", value = "STATCAL Likert Scale Distribution Analysis Report")),
                column(6, textInput("word_report_subtitle", "Word report subtitle", value = "Frequency, percentage, split distribution, and stacked bar chart"))
              ),
              fluidRow(
                column(6, textInput("table1_title", "Table 1 title", value = "Table 1. Likert Frequency and Percentage Distribution")),
                column(6, textInput("table2_title", "Table 2 title", value = "Table 2. Likert Frequency and Percentage Distribution by Split Variable"))
              ),
              fluidRow(
                column(6, textInput("table3_title", "Table 3 title", value = "Table 3. Data Used to Create the Stacked Bar Chart")),
                column(6, textInput("figure1_title", "Figure 1 title", value = "Figure 1. Likert Scale Distribution Stacked Bar Chart"))
              ),
              checkboxInput(
                "word_include_interpretation",
                "Include automatic interpretation below each result table and figure",
                value = TRUE
              ),
              selectInput(
                "word_interpretation_language",
                "Automatic interpretation language",
                choices = c("Bahasa Indonesia", "English"),
                selected = "Bahasa Indonesia"
              ),
              checkboxInput("word_include_chart_data", "Include Chart Data table in Word", value = TRUE),
              checkboxInput("word_include_raw_data", "Include Raw Data appendix in Word", value = FALSE),
              tags$p(
                class = "small-note",
                "Automatic interpretation is descriptive. It summarizes the observed Likert distributions and does not imply statistical significance or causal relationships."
              ),
              tags$p(
                class = "small-note",
                paste0(
                  "Generated files are created and verified inside the webR session, ",
                  "then delivered directly to the browser as binary Blob downloads. ",
                  "This avoids Shinylive/GitHub Pages Service Worker download errors."
                )
              ))
        ),
        fluidRow(
          box(width = 4, title = "Stacked Bar Chart PNG", status = "warning", solidHeader = TRUE,
              actionButton("generate_chart_png", "Generate PNG", icon = icon("image")),
              br(), br(), uiOutput("chart_generated_download_ui"),
              br(), actionButton("download_chart_fallback", "Generate & Download PNG", icon = icon("download"))),
          box(width = 4, title = "Analysis Excel", status = "success", solidHeader = TRUE,
              actionButton("generate_excel", "Generate Analysis Excel", icon = icon("file-excel")),
              br(), br(), uiOutput("excel_generated_download_ui"),
              br(), actionButton("download_excel_fallback", "Generate & Download Excel", icon = icon("download"))),
          box(width = 4, title = "Analysis Word", status = "info", solidHeader = TRUE,
              actionButton("generate_word", "Generate Analysis Word", icon = icon("file-word")),
              br(), br(), uiOutput("word_generated_download_ui"),
              br(), actionButton("download_word_fallback", "Generate & Download Word", icon = icon("download")))
        )
      )
    )
  )
)

# ============================================================
# SERVER
# ============================================================

server <- function(input, output, session) {
  chart_export_result <- reactiveVal(NULL)
  excel_export_result <- reactiveVal(NULL)
  word_export_result <- reactiveVal(NULL)
  session_export_dir <- create_session_export_dir(session$token)
  
  session$onSessionEnded(function() {
    try(unlink(session_export_dir, recursive = TRUE, force = TRUE), silent = TRUE)
  })
  
  output$chart_generated_download_ui <- renderUI({
    generated_download_ui(chart_export_result(), "download_chart_generated", "Download Generated PNG", "image")
  })
  output$excel_generated_download_ui <- renderUI({
    generated_download_ui(excel_export_result(), "download_excel_generated", "Download Generated Excel", "file-excel")
  })
  output$word_generated_download_ui <- renderUI({
    generated_download_ui(word_export_result(), "download_word_generated", "Download Generated Word", "file-word")
  })
  
  current_excel_path <- reactive({
    if (!is.null(input$uploaded_file)) {
      input$uploaded_file$datapath
    } else if (file.exists(SAMPLE_DATA_PATH)) {
      SAMPLE_DATA_PATH
    } else {
      NULL
    }
  })
  
  output$sheet_ui <- renderUI({
    path <- current_excel_path()
    if (is.null(path)) return(helpText("Please upload an Excel file to start the analysis."))
    sheets <- readxl::excel_sheets(path)
    selectInput("sheet_name", "Worksheet", choices = sheets, selected = sheets[1])
  })
  
  data_raw <- reactive({
    path <- current_excel_path()
    req(path)
    sheets <- readxl::excel_sheets(path)
    sheet <- input$sheet_name
    if (is.null(sheet) || !(sheet %in% sheets)) sheet <- sheets[1]
    clean_dataframe(readxl::read_excel(path, sheet = sheet))
  })
  
  likert_levels <- reactive({
    parse_likert_levels(input$likert_order_text)
  })
  
  likely_item_columns <- reactive({
    detect_likert_item_columns(data_raw(), likert_levels())
  })
  
  selected_items <- reactive({
    vals <- input$item_cols
    if (is.null(vals) || length(vals) == 0) vals <- likely_item_columns()
    vals[vals %in% names(data_raw())]
  })
  
  output$item_vars_ui <- renderUI({
    df <- data_raw()
    defaults <- likely_item_columns()
    if (length(defaults) == 0) defaults <- names(df)
    selectizeInput("item_cols", "Select item question variables", choices = names(df), selected = defaults, multiple = TRUE,
                   options = list(plugins = list("remove_button")))
  })
  
  output$split_var_ui <- renderUI({
    df <- data_raw()
    choices <- c("None", setdiff(names(df), selected_items()))
    current <- isolate(input$split_var)
    preferred <- if (!is.null(current) && current %in% choices) current else if ("Jenis Kelamin" %in% choices) "Jenis Kelamin" else if ("Pendidikan" %in% choices) "Pendidikan" else "None"
    selectInput("split_var", "Split by one categorical variable", choices = choices, selected = preferred)
  })
  
  output$split_order_ui <- renderUI({
    df <- data_raw()
    split_col <- input$split_var
    if (is.null(split_col) || split_col == "None" || !(split_col %in% names(df))) {
      return(tags$p(class = "small-note", "No split order is needed when split variable is None."))
    }
    split_values <- sorted_unique_values(as.character(df[[split_col]]))
    textInput("split_order_text", "Split category order for table and panels (comma-separated)", value = paste(split_values, collapse = ", "))
  })
  
  output$palette_preview_ui <- renderUI({
    levels <- likert_levels()
    palette_name <- input$chart_palette
    if (is.null(palette_name)) palette_name <- "Scopus-inspired Orange–Blue"
    cols <- if (palette_name == "Manual / Custom") get_manual_colors(levels, input) else {
      tmp <- palette_colors(palette_name, length(levels), isTRUE(input$chart_reverse_palette)); names(tmp) <- levels; tmp
    }
    tags$div(
      style = "display:flex;gap:5px;flex-wrap:wrap;margin-bottom:10px;",
      lapply(seq_along(levels), function(i) {
        tags$div(title = paste(levels[i], cols[i]), style = paste0("width:34px;height:20px;border-radius:4px;border:1px solid #aaa;background:", cols[i], ";"))
      })
    )
  })
  
  output$color_settings_ui <- renderUI({
    levels <- likert_levels()
    if (!identical(input$chart_palette, "Manual / Custom")) {
      return(tags$p(class = "small-note", "Manual color pickers are inactive because a predefined palette is selected."))
    }
    defaults <- palette_colors("Scopus-inspired Orange–Blue", length(levels))
    tagList(lapply(seq_along(levels), function(i) {
      colourpicker::colourInput(color_input_id(levels[i]), paste0("Color for ", levels[i]), value = defaults[i], showColour = "both")
    }))
  })
  
  output$legend_label_settings_ui <- renderUI({
    levels <- likert_levels()
    tagList(lapply(levels, function(lv) {
      textInput(legend_input_id(lv), paste0("Legend label for ", lv), value = default_legend_label(lv))
    }))
  })
  
  output$metric_rows <- renderValueBox({ valueBox(nrow(data_raw()), "Rows", icon = icon("table"), color = "blue") })
  output$metric_columns <- renderValueBox({ valueBox(ncol(data_raw()), "Columns", icon = icon("columns"), color = "yellow") })
  output$metric_items <- renderValueBox({ valueBox(length(selected_items()), "Selected item questions", icon = icon("list"), color = "green") })
  output$metric_levels <- renderValueBox({ valueBox(length(likert_levels()), "Likert labels", icon = icon("sort"), color = "purple") })
  
  output$data_preview <- renderDT({
    DT::datatable(make_display_safe(data_raw()), options = list(scrollX = TRUE, pageLength = 10))
  })
  
  frequency_chart_data <- reactive({
    build_likert_long_data(data_raw(), selected_items(), likert_levels(), split_col = "None", digits = input$decimal_digits)
  })
  
  chart_data <- reactive({
    build_likert_long_data(data_raw(), selected_items(), likert_levels(), input$split_var, input$decimal_digits, input$split_order_text)
  })
  
  frequency_table_data <- reactive({
    format_likert_wide_table_from_long(frequency_chart_data(), likert_levels(), input$decimal_digits, include_split = FALSE)
  })
  
  split_table_data <- reactive({
    if (is.null(input$split_var) || input$split_var == "None" || !(input$split_var %in% names(data_raw()))) return(frequency_table_data())
    format_likert_wide_table_from_long(chart_data(), likert_levels(), input$decimal_digits, split_col = input$split_var, include_split = TRUE)
  })
  
  stacked_bar_plot_object <- reactive({
    create_stacked_bar_plot(
      chart_df = chart_data(),
      levels = likert_levels(),
      manual_colors = get_chart_colors(likert_levels(), input),
      legend_labels = get_legend_labels(likert_levels(), input),
      legend_title = input$chart_legend_title,
      title = input$chart_title,
      subtitle = input$chart_subtitle,
      x_axis_title = input$chart_x_axis_title,
      y_axis_title = input$chart_y_axis_title,
      y_metric = input$chart_y_metric,
      label_mode = input$chart_label_mode,
      orientation = input$chart_orientation,
      theme_name = input$chart_theme,
      bar_width = input$bar_width,
      show_labels = input$chart_show_labels,
      label_size = input$chart_label_size,
      label_color = input$chart_label_color,
      label_vjust = input$chart_label_vjust,
      facet_by_split = input$chart_facet_split,
      split_panel_cols = input$chart_split_cols,
      facet_scales = input$chart_facet_scales,
      legend_position = input$chart_legend_position,
      title_size = input$chart_title_size,
      subtitle_size = input$chart_subtitle_size,
      axis_title_size = input$chart_axis_title_size,
      axis_text_size = input$chart_axis_text_size,
      legend_title_size = input$chart_legend_title_size,
      legend_text_size = input$chart_legend_text_size,
      x_text_angle = input$chart_x_text_angle,
      digits = input$decimal_digits
    )
  })
  
  output$frequency_table <- renderDT({
    DT::datatable(make_display_safe(frequency_table_data()), options = list(scrollX = TRUE, pageLength = 15))
  })
  output$split_frequency_table <- renderDT({
    DT::datatable(make_display_safe(split_table_data()), options = list(scrollX = TRUE, pageLength = 15))
  })
  output$stacked_plot_ui <- renderUI({
    plotOutput("stacked_bar_plot", height = paste0(safe_number(input$chart_height_px, 700, 350, 1400), "px"))
  })
  output$stacked_bar_plot <- renderPlot({ stacked_bar_plot_object() })
  output$chart_data_table <- renderDT({
    DT::datatable(make_display_safe(chart_data()), options = list(scrollX = TRUE, pageLength = 15))
  })
  
  build_export_metadata <- reactive({
    split_order_value <- if (is.null(input$split_order_text)) "" else as.character(input$split_order_text)
    split_var_value <- if (is.null(input$split_var)) "None" else as.character(input$split_var)
    data.frame(
      Item = c(
        "Application", "Export Time", "Rows", "Selected Item Questions", "Likert Order",
        "Legend Labels", "Color Palette", "Decimal Digits", "Split Variable", "Split Category Order",
        "Chart Bar Height", "Chart Label Mode"
      ),
      Value = c(
        APP_TITLE, as.character(Sys.time()), as.character(nrow(data_raw())),
        paste(selected_items(), collapse = ", "), paste(likert_levels(), collapse = ", "),
        paste(paste(likert_levels(), get_legend_labels(likert_levels(), input), sep = " = "), collapse = "; "),
        input$chart_palette, as.character(input$decimal_digits), split_var_value, split_order_value,
        input$chart_y_metric, input$chart_label_mode
      ),
      stringsAsFactors = FALSE
    )
  })
  
  export_current_excel <- function(file) {
    export_likert_workbook(
      file = file,
      raw_df = data_raw(),
      frequency_df = frequency_table_data(),
      split_df = split_table_data(),
      chart_df = chart_data(),
      metadata_df = build_export_metadata(),
      plot_object = stacked_bar_plot_object(),
      chart_title = input$figure1_title,
      table1_title = input$table1_title,
      table2_title = input$table2_title,
      table3_title = input$table3_title,
      export_width = input$export_width,
      export_height = input$export_height
    )
  }
  
  word_narratives <- reactive({
    build_likert_word_narratives(
      overall_long_df = frequency_chart_data(),
      split_long_df = chart_data(),
      levels = likert_levels(),
      legend_labels = get_legend_labels(likert_levels(), input),
      split_var = ifelse(is.null(input$split_var), "None", input$split_var),
      digits = input$decimal_digits,
      chart_metric = input$chart_y_metric,
      facet_by_split = isTRUE(input$chart_facet_split),
      language = input$word_interpretation_language
    )
  })
  
  export_current_word <- function(file) {
    word_width <- min(safe_number(input$export_width, 8, 4, 12), 6.5)
    ratio <- safe_number(input$export_height, 6, 3, 20) / safe_number(input$export_width, 8, 4, 30)
    word_height <- max(3, min(8.5, word_width * ratio))
    export_word_report(
      file = file,
      report_title = input$word_report_title,
      report_subtitle = input$word_report_subtitle,
      metadata_df = build_export_metadata(),
      frequency_df = frequency_table_data(),
      split_df = split_table_data(),
      chart_df = chart_data(),
      raw_df = data_raw(),
      plot_object = stacked_bar_plot_object(),
      table1_title = input$table1_title,
      table2_title = input$table2_title,
      table3_title = input$table3_title,
      figure1_title = input$figure1_title,
      include_chart_data = isTRUE(input$word_include_chart_data),
      include_raw_data = isTRUE(input$word_include_raw_data),
      include_interpretation = isTRUE(input$word_include_interpretation),
      table1_narrative = word_narratives()$table1,
      table2_narrative = word_narratives()$table2,
      figure1_narrative = word_narratives()$figure1,
      table3_narrative = word_narratives()$table3,
      plot_width = word_width,
      plot_height = word_height
    )
  }
  
  make_result_safe <- function(expr, filename, success_message) {
    tryCatch({
      path <- file.path(session_export_dir, filename)
      expr(path)
      new_export_result(path, filename, success_message)
    }, error = function(e) {
      list(ok = FALSE, message = paste("Export failed:", conditionMessage(e)), file = NULL, filename = filename, size = 0)
    })
  }
  
  observeEvent(input$generate_chart_png, {
    filename <- make_export_filename("statcal_likert_stacked_bar_chart", "png", input$export_dpi)
    result <- make_result_safe(
      function(path) generate_plot_png(stacked_bar_plot_object(), path, input$export_width, input$export_height, input$export_dpi, safe_theme_bg(input$chart_theme)),
      filename,
      "PNG figure has been generated and verified successfully."
    )
    chart_export_result(result)
  })
  
  observeEvent(input$generate_excel, {
    filename <- make_export_filename("statcal_likert_analysis", "xlsx")
    result <- make_result_safe(function(path) export_current_excel(path), filename, "Excel workbook has been generated and verified successfully.")
    excel_export_result(result)
  })
  
  observeEvent(input$generate_word, {
    filename <- make_export_filename("statcal_likert_analysis_report", "docx")
    result <- make_result_safe(function(path) export_current_word(path), filename, "Word report with automatic interpretation has been generated and verified successfully.")
    word_export_result(result)
  })
  
  # ----------------------------------------------------------
  # BROWSER BLOB DOWNLOADS (SHINYLIVE / GITHUB PAGES)
  # ----------------------------------------------------------
  # Generated files live inside webR's virtual filesystem.
  # They are encoded as Base64, sent to JavaScript, reconstructed as a
  # binary Blob, and downloaded by the browser. No downloadHandler URL
  # or Service Worker download route is required.
  
  send_result_safely <- function(result, mime) {
    tryCatch({
      req(
        result,
        isTRUE(result$ok),
        !is.null(result$file),
        file.exists(result$file)
      )
      
      send_binary_file_to_browser(
        session = session,
        path = result$file,
        filename = result$filename,
        mime = mime
      )
      
    }, error = function(e) {
      showNotification(
        paste("Download failed:", conditionMessage(e)),
        type = "error",
        duration = NULL
      )
    })
  }
  
  # Download a file that has already been generated and verified.
  observeEvent(input$download_chart_generated, {
    send_result_safely(
      chart_export_result(),
      "image/png"
    )
  }, ignoreInit = TRUE)
  
  observeEvent(input$download_excel_generated, {
    send_result_safely(
      excel_export_result(),
      "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )
  }, ignoreInit = TRUE)
  
  observeEvent(input$download_word_generated, {
    send_result_safely(
      word_export_result(),
      "application/vnd.openxmlformats-officedocument.wordprocessingml.document"
    )
  }, ignoreInit = TRUE)
  
  # ----------------------------------------------------------
  # ONE-CLICK GENERATE + DOWNLOAD
  # ----------------------------------------------------------
  
  observeEvent(input$download_chart_fallback, {
    filename <- make_export_filename(
      "statcal_likert_stacked_bar_chart",
      "png",
      input$export_dpi
    )
    
    result <- make_result_safe(
      function(path) {
        generate_plot_png(
          stacked_bar_plot_object(),
          path,
          input$export_width,
          input$export_height,
          input$export_dpi,
          safe_theme_bg(input$chart_theme)
        )
      },
      filename,
      "PNG figure has been generated and verified successfully."
    )
    
    chart_export_result(result)
    
    if (isTRUE(result$ok)) {
      send_result_safely(result, "image/png")
    } else {
      showNotification(
        result$message,
        type = "error",
        duration = NULL
      )
    }
  }, ignoreInit = TRUE)
  
  observeEvent(input$download_excel_fallback, {
    filename <- make_export_filename(
      "statcal_likert_analysis",
      "xlsx"
    )
    
    result <- make_result_safe(
      function(path) export_current_excel(path),
      filename,
      "Excel workbook has been generated and verified successfully."
    )
    
    excel_export_result(result)
    
    if (isTRUE(result$ok)) {
      send_result_safely(
        result,
        "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
      )
    } else {
      showNotification(
        result$message,
        type = "error",
        duration = NULL
      )
    }
  }, ignoreInit = TRUE)
  
  observeEvent(input$download_word_fallback, {
    filename <- make_export_filename(
      "statcal_likert_analysis_report",
      "docx"
    )
    
    result <- make_result_safe(
      function(path) export_current_word(path),
      filename,
      "Word report with automatic interpretation has been generated and verified successfully."
    )
    
    word_export_result(result)
    
    if (isTRUE(result$ok)) {
      send_result_safely(
        result,
        "application/vnd.openxmlformats-officedocument.wordprocessingml.document"
      )
    } else {
      showNotification(
        result$message,
        type = "error",
        duration = NULL
      )
    }
  }, ignoreInit = TRUE)
}

shinyApp(ui = ui, server = server)
