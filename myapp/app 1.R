# ============================================================
# STATCAL ONLINE - Likert Scale Distribution and Stacked Bar Chart Visualizer
# R Shiny Version
# ============================================================
# Required packages:
# install.packages(c(
#   "shiny", "shinydashboard", "DT", "readxl", "dplyr", "tidyr",
#   "ggplot2", "shinycssloaders", "scales", "openxlsx"
# ))

required_packages <- c(
  "shiny", "shinydashboard", "DT", "readxl", "dplyr", "tidyr",
  "ggplot2", "shinycssloaders", "scales", "openxlsx"
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

# ============================================================
# CONSTANTS
# ============================================================

APP_NAME <- "STATCAL ONLINE"
APP_TITLE <- "Likert Scale Distribution and Stacked Bar Chart Visualizer"
APP_UPDATED <- "Last updated on June 19, 2026"
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

safe_id <- function(x) {
  x <- gsub("[^A-Za-z0-9_]", "_", as.character(x))
  x <- gsub("_+", "_", x)
  paste0("color_", x)
}

is_hex_color <- function(x) {
  is.character(x) && length(x) == 1 && grepl("^#([0-9A-Fa-f]{6}|[0-9A-Fa-f]{8})$", x)
}

get_manual_colors <- function(levels, input) {
  out <- character(length(levels))
  for (i in seq_along(levels)) {
    default_col <- DEFAULT_COLORS[((i - 1) %% length(DEFAULT_COLORS)) + 1]
    val <- input[[safe_id(levels[i])]]
    if (is.null(val) || length(val) == 0 || !is_hex_color(val)) val <- default_col
    out[i] <- val
  }
  names(out) <- levels
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
                                    title, subtitle,
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
  y_label <- if (y_metric == "Frequency") "Frequency" else "Percentage (%)"
  p <- ggplot(chart_df, aes(x = `Item Pertanyaan`, y = Y_Value, fill = Response)) +
    geom_col(width = bar_width, color = "white", linewidth = 0.25) +
    scale_fill_manual(values = manual_colors, drop = FALSE) +
    labs(title = title, subtitle = subtitle, x = "Item Pertanyaan", y = y_label, fill = "Likert Response") +
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
  if (facet_by_split && any(chart_df$Split != "All")) {
    facet_scales <- ifelse(is.null(facet_scales) || !(facet_scales %in% c("fixed", "free_y", "free_x", "free")), "fixed", facet_scales)
    p <- p + facet_wrap(~ Split, ncol = split_panel_cols, scales = facet_scales)
  }
  if (orientation == "Horizontal") {
    p <- p + coord_flip()
  }
  p + theme(plot.background = element_rect(fill = th$figure_facecolor, color = NA))
}

# ============================================================
# SAFE STATIC EXPORT HELPERS
# ============================================================

write_error_png <- function(file, message, width = 8, height = 5, dpi = 300, bg = "white") {
  width <- safe_number(width, 8, 3, 20)
  height <- safe_number(height, 5, 3, 20)
  dpi <- safe_number(dpi, 300, 72, 1500)
  grDevices::png(filename = file, width = width, height = height, units = "in", res = dpi, bg = bg)
  on.exit(grDevices::dev.off(), add = TRUE)
  graphics::par(bg = bg, mar = c(1, 1, 1, 1))
  graphics::plot.new()
  graphics::text(0.5, 0.62, "STATCAL ONLINE", cex = 1.6, font = 2)
  graphics::text(0.5, 0.50, "PNG export could not be generated.", cex = 1.1)
  graphics::text(0.5, 0.40, paste("Reason:", message), cex = 0.8)
}

make_static_export_dir <- function() {
  export_dir <- file.path(getwd(), "www", "statcal_exports")
  if (!dir.exists(export_dir)) {
    dir.create(export_dir, recursive = TRUE, showWarnings = FALSE)
  }
  export_dir
}

make_static_export_dir()

make_export_filename <- function(prefix, ext = "png", dpi = NULL) {
  stamp <- format(Sys.time(), "%Y%m%d_%H%M%S")
  if (!is.null(dpi)) {
    paste0(prefix, "_", dpi, "dpi_", stamp, ".", ext)
  } else {
    paste0(prefix, "_", stamp, ".", ext)
  }
}

make_static_href <- function(filename) {
  paste0("statcal_exports/", utils::URLencode(filename, reserved = TRUE))
}

export_ggplot_static_png <- function(plot_function, prefix, width, height, dpi, bg = "white") {
  export_dir <- make_static_export_dir()
  width <- safe_number(width, 8, 3, 30)
  height <- safe_number(height, 6, 3, 30)
  dpi <- safe_number(dpi, 1200, 72, 1500)
  bg <- ifelse(is.null(bg) || length(bg) == 0 || is.na(bg), "white", bg)
  filename <- make_export_filename(prefix, "png", dpi)
  out_file <- file.path(export_dir, filename)
  result <- tryCatch({
    plot_object <- plot_function()
    if (!inherits(plot_object, "ggplot")) {
      stop("The selected output is not a ggplot object. Please check the selected variables.")
    }
    grDevices::png(
      filename = out_file,
      width = width,
      height = height,
      units = "in",
      res = dpi,
      bg = bg,
      type = ifelse(.Platform$OS.type == "windows", "windows", "cairo")
    )
    print(plot_object)
    grDevices::dev.off()
    if (!file.exists(out_file) || file.info(out_file)$size <= 0) {
      stop("The PNG file was not created.")
    }
    list(
      ok = TRUE,
      message = "PNG file has been generated successfully.",
      file = normalizePath(out_file, winslash = "/", mustWork = FALSE),
      href = make_static_href(filename),
      filename = filename
    )
  }, error = function(e) {
    while (grDevices::dev.cur() > 1) {
      try(grDevices::dev.off(), silent = TRUE)
    }
    write_error_png(out_file, conditionMessage(e), width = width, height = height, dpi = min(dpi, 600), bg = bg)
    list(
      ok = FALSE,
      message = paste("PNG export failed, but an error PNG was generated:", conditionMessage(e)),
      file = normalizePath(out_file, winslash = "/", mustWork = FALSE),
      href = make_static_href(filename),
      filename = filename
    )
  })
  result
}

static_export_link_ui <- function(result, button_label = "Download generated file", open_label = "Open file in new tab", preview_image = FALSE) {
  if (is.null(result)) {
    return(tags$p("Click Generate first, then the download link will appear here."))
  }
  ui <- tagList(
    tags$p(if (isTRUE(result$ok)) result$message else result$message),
    tags$a(
      href = result$href,
      download = result$filename,
      target = "_blank",
      class = "btn btn-success",
      icon("download"),
      button_label
    ),
    tags$span(" "),
    tags$a(
      href = result$href,
      target = "_blank",
      class = "btn btn-info",
      icon("external-link-alt"),
      open_label
    ),
    tags$p(style = "font-size: 12px; margin-top: 8px; color: #555;", paste("Generated file:", result$filename)),
    tags$p(style = "font-size: 11px; color: #777; word-break: break-all;", paste("Local file path:", result$file))
  )
  if (preview_image) {
    ui <- tagList(ui, tags$img(src = result$href, style = "max-width: 100%; margin-top: 8px; border: 1px solid #ddd;"))
  }
  ui
}

# ============================================================
# EXCEL EXPORT HELPERS
# ============================================================

safe_sheet_name <- function(name) {
  name <- as.character(name)
  invalid_chars <- c("\\", "/", "?", "*", "[", "]", ":")
  for (ch in invalid_chars) {
    name <- gsub(ch, "_", name, fixed = TRUE)
  }
  name <- trimws(name)
  name <- substr(name, 1, 31)
  ifelse(nchar(name) == 0, "Sheet", name)
}

write_table_sheet <- function(wb, sheet_name, df) {
  sheet_name <- safe_sheet_name(sheet_name)
  openxlsx::addWorksheet(wb, sheet_name)
  df <- make_display_safe(as.data.frame(df))
  openxlsx::writeData(wb, sheet_name, df)
  if (ncol(df) > 0) {
    openxlsx::setColWidths(wb, sheet_name, cols = 1:ncol(df), widths = "auto")
    header_style <- openxlsx::createStyle(textDecoration = "bold", fgFill = "#D9EAF7", border = "Bottom")
    openxlsx::addStyle(wb, sheet_name, header_style, rows = 1, cols = 1:ncol(df), gridExpand = TRUE)
    openxlsx::freezePane(wb, sheet_name, firstRow = TRUE)
  }
}

export_likert_workbook <- function(file, raw_df, frequency_df, split_df, chart_df, metadata_df) {
  wb <- openxlsx::createWorkbook()
  write_table_sheet(wb, "Export Info", metadata_df)
  write_table_sheet(wb, "Raw Data", raw_df)
  write_table_sheet(wb, "Frequency Table", frequency_df)
  write_table_sheet(wb, "Split Frequency Table", split_df)
  write_table_sheet(wb, "Chart Data", chart_df)
  openxlsx::saveWorkbook(wb, file, overwrite = TRUE)
}

export_excel_static_file <- function(export_function, prefix) {
  export_dir <- make_static_export_dir()
  filename <- make_export_filename(prefix, "xlsx")
  out_file <- file.path(export_dir, filename)
  result <- tryCatch({
    export_function(out_file)
    if (!file.exists(out_file) || file.info(out_file)$size <= 0) {
      stop("The Excel file was not created.")
    }
    list(
      ok = TRUE,
      message = "Excel file has been generated successfully.",
      file = normalizePath(out_file, winslash = "/", mustWork = FALSE),
      href = make_static_href(filename),
      filename = filename
    )
  }, error = function(e) {
    list(
      ok = FALSE,
      message = paste("Excel export failed:", conditionMessage(e)),
      file = normalizePath(out_file, winslash = "/", mustWork = FALSE),
      href = make_static_href(filename),
      filename = filename
    )
  })
  result
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
      tags$style(HTML("\n        .content-wrapper, .right-side { background-color: #f7f9fb; }\n        .box { border-radius: 10px; }\n        .statcal-title { font-size: 24px; font-weight: 700; color: #1F4E79; }\n        .statcal-subtitle { font-size: 18px; font-weight: 600; color: #333333; }\n        .statcal-note { line-height: 1.6; text-align: justify; }\n        .small-note { font-size: 12px; color: #666666; }\n      "))
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
          box(width = 3, title = "Bar Style", status = "primary", solidHeader = TRUE,
              selectInput("chart_theme", "Background theme", choices = names(THEMES), selected = "White Publication"),
              sliderInput("bar_width", "Bar width", min = 0.25, max = 1.00, value = 0.72, step = 0.01),
              sliderInput("chart_height_px", "Preview chart height (px)", min = 350, max = 1200, value = 700, step = 50),
              selectInput("chart_legend_position", "Legend position", choices = legend_choices, selected = "Right"),
              sliderInput("chart_x_text_angle", "X-axis text angle", min = 0, max = 90, value = 35, step = 5)),
          box(width = 3, title = "Text and Labels", status = "primary", solidHeader = TRUE,
              textInput("chart_title", "Title", value = "Likert Scale Distribution"),
              textInput("chart_subtitle", "Subtitle", value = "Stacked bar chart based on selected survey items"),
              checkboxInput("chart_show_labels", "Show text labels", value = TRUE),
              sliderInput("chart_label_size", "Label text size", min = 2, max = 8, value = 3.2, step = 0.2),
              textInput("chart_label_color", "Label text color", value = "#111111"),
              sliderInput("chart_label_vjust", "Label position inside stack", min = 0.05, max = 0.95, value = 0.50, step = 0.05)),
          box(width = 3, title = "Manual Bar Colors", status = "primary", solidHeader = TRUE,
              tags$p(class = "small-note", "Use HEX color codes, for example #2166AC."),
              uiOutput("color_settings_ui"))
        ),
        fluidRow(
          box(width = 12, title = "Flexible Text Size Settings", status = "info", solidHeader = TRUE, collapsible = TRUE, collapsed = TRUE,
              fluidRow(
                column(2, sliderInput("chart_title_size", "Title", 8, 34, 16, 1)),
                column(2, sliderInput("chart_subtitle_size", "Subtitle", 6, 26, 11, 1)),
                column(2, sliderInput("chart_axis_title_size", "Axis title", 6, 24, 11, 1)),
                column(2, sliderInput("chart_axis_text_size", "Axis text", 5, 22, 9, 1)),
                column(2, sliderInput("chart_legend_title_size", "Legend title", 5, 24, 10, 1)),
                column(2, sliderInput("chart_legend_text_size", "Legend text", 5, 22, 9, 1))
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
        "5. Export",
        br(),
        fluidRow(
          box(width = 12, title = "Export Settings", status = "primary", solidHeader = TRUE,
              fluidRow(
                column(4, selectInput("export_dpi", "PNG resolution / DPI", choices = c(300, 600, 900, 1200, 1500), selected = 1200)),
                column(4, numericInput("export_width", "Export width (inches)", value = 8, min = 4, max = 30, step = 0.5)),
                column(4, numericInput("export_height", "Export height (inches)", value = 6, min = 3, max = 30, step = 0.5))
              ),
              tags$p(tags$b("Default DPI: "), "1200 DPI for publication-ready output."))
        ),
        fluidRow(
          box(width = 6, title = "Stacked Bar Chart PNG", status = "warning", solidHeader = TRUE,
              actionButton("generate_chart_png", "Generate Stacked Bar Chart PNG", icon = icon("image")),
              br(), br(), uiOutput("chart_static_download_ui")),
          box(width = 6, title = "Likert Tables Excel", status = "success", solidHeader = TRUE,
              actionButton("generate_excel", "Generate Likert Tables Excel", icon = icon("file-excel")),
              br(), br(), uiOutput("excel_static_download_ui"),
              br(),
              downloadButton("download_excel_fallback", "Fallback Download Excel"))
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
  
  output$chart_static_download_ui <- renderUI({
    static_export_link_ui(chart_export_result(), "Download Stacked Bar Chart PNG", "Open PNG in new tab", preview_image = TRUE)
  })
  
  output$excel_static_download_ui <- renderUI({
    static_export_link_ui(excel_export_result(), "Download Likert Tables Excel", "Open Excel file in new tab", preview_image = FALSE)
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
    if (is.null(path)) {
      return(helpText("Please upload an Excel file to start the analysis."))
    }
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
  
  likely_item_columns <- reactive({
    df <- data_raw()
    cols <- names(df)
    defaults <- grep("pertanyaan|question|item", cols, ignore.case = TRUE, value = TRUE)
    if (length(defaults) == 0) defaults <- cols
    defaults
  })
  
  selected_items <- reactive({
    if (is.null(input$item_cols) || length(input$item_cols) == 0) {
      likely_item_columns()
    } else {
      input$item_cols
    }
  })
  
  likert_levels <- reactive({
    parse_likert_levels(input$likert_order_text)
  })
  
  output$item_vars_ui <- renderUI({
    df <- data_raw()
    selectizeInput("item_cols", "Select item question variables", choices = names(df), selected = likely_item_columns(), multiple = TRUE)
  })
  
  output$split_var_ui <- renderUI({
    df <- data_raw()
    choices <- c("None", setdiff(names(df), selected_items()))
    selected <- if ("Jenis Kelamin" %in% choices) "Jenis Kelamin" else if ("Pendidikan" %in% choices) "Pendidikan" else "None"
    selectInput("split_var", "Split by one categorical variable", choices = choices, selected = selected)
  })
  
  output$split_order_ui <- renderUI({
    df <- data_raw()
    split_col <- input$split_var
    if (is.null(split_col) || split_col == "None" || !(split_col %in% names(df))) {
      return(tags$p(class = "small-note", "No split order is needed when split variable is None."))
    }
    split_values <- sorted_unique_values(as.character(df[[split_col]]))
    textInput(
      "split_order_text",
      "Split category order for table and panels (comma-separated)",
      value = paste(split_values, collapse = ", ")
    )
  })
  
  output$color_settings_ui <- renderUI({
    levels <- likert_levels()
    tagList(lapply(seq_along(levels), function(i) {
      textInput(safe_id(levels[i]), paste0("Color for ", levels[i]), value = DEFAULT_COLORS[((i - 1) %% length(DEFAULT_COLORS)) + 1])
    }))
  })
  
  output$metric_rows <- renderValueBox({
    valueBox(nrow(data_raw()), "Rows", icon = icon("table"), color = "blue")
  })
  output$metric_columns <- renderValueBox({
    valueBox(ncol(data_raw()), "Columns", icon = icon("columns"), color = "yellow")
  })
  output$metric_items <- renderValueBox({
    valueBox(length(selected_items()), "Selected item questions", icon = icon("list"), color = "green")
  })
  output$metric_levels <- renderValueBox({
    valueBox(length(likert_levels()), "Likert labels", icon = icon("sort"), color = "purple")
  })
  
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
    if (is.null(input$split_var) || input$split_var == "None" || !(input$split_var %in% names(data_raw()))) {
      return(frequency_table_data())
    }
    format_likert_wide_table_from_long(chart_data(), likert_levels(), input$decimal_digits, split_col = input$split_var, include_split = TRUE)
  })
  
  stacked_bar_plot_object <- reactive({
    create_stacked_bar_plot(
      chart_df = chart_data(),
      levels = likert_levels(),
      manual_colors = get_manual_colors(likert_levels(), input),
      title = input$chart_title,
      subtitle = input$chart_subtitle,
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
  
  output$stacked_bar_plot <- renderPlot({
    stacked_bar_plot_object()
  })
  
  output$chart_data_table <- renderDT({
    DT::datatable(make_display_safe(chart_data()), options = list(scrollX = TRUE, pageLength = 15))
  })
  
  build_export_metadata <- reactive({
    split_order_value <- if (is.null(input$split_order_text)) "" else as.character(input$split_order_text)
    split_var_value <- if (is.null(input$split_var)) "None" else as.character(input$split_var)
    data.frame(
      Item = c(
        "Application", "Export Time", "Rows", "Selected Item Questions", "Likert Order",
        "Decimal Digits", "Split Variable", "Split Category Order", "Chart Bar Height", "Chart Label Mode"
      ),
      Value = c(
        APP_TITLE, as.character(Sys.time()), as.character(nrow(data_raw())),
        paste(selected_items(), collapse = ", "), paste(likert_levels(), collapse = ", "),
        as.character(input$decimal_digits), split_var_value, split_order_value, input$chart_y_metric, input$chart_label_mode
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
      metadata_df = build_export_metadata()
    )
  }
  
  observeEvent(input$generate_chart_png, {
    result <- export_ggplot_static_png(
      plot_function = function() stacked_bar_plot_object(),
      prefix = "statcal_online_likert_stacked_bar_chart",
      width = input$export_width,
      height = input$export_height,
      dpi = input$export_dpi,
      bg = safe_theme_bg(input$chart_theme)
    )
    chart_export_result(result)
  })
  
  observeEvent(input$generate_excel, {
    result <- export_excel_static_file(
      export_function = function(path) export_current_excel(path),
      prefix = "statcal_online_likert_frequency_tables"
    )
    excel_export_result(result)
  })
  
  output$download_excel_fallback <- downloadHandler(
    filename = function() paste0("statcal_online_likert_frequency_tables_", format(Sys.time(), "%Y%m%d_%H%M%S"), ".xlsx"),
    contentType = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    content = function(file) {
      export_current_excel(file)
    }
  )
}

shinyApp(ui = ui, server = server)
