# =========================================================
# Our Changing Travel — ≥ weekly SRN use (faceted blocks)
# =========================================================

library(readxl)
library(dplyr)
library(ggplot2)
library(stringr)
library(tibble)
library(grid)
library(gtable)
# ---------- utilities ----------
numify <- function(x) as.numeric(gsub("[^0-9.\\-]", "", as.character(x)))
clean_names <- function(v) gsub("\\.\\.\\.[0-9]+$", "", v)

xl_col_index <- function(x) {
  sapply(x, function(s) {
    s <- toupper(gsub("\\s+", "", s))
    ch <- strsplit(s, "")[[1]]
    sum((match(ch, LETTERS)) * 26^(rev(seq_along(ch)) - 1))
  })
}

read_oct <- function(path, sheet) {
  df <- read_xlsx(path, sheet = sheet, skip = 10)
  names(df) <- clean_names(names(df))
  df
}
add_facet_separators <- function(p, colour = "grey60", linetype = "dashed", size = 0.6){
  g <- ggplotGrob(p)
  panels <- grep("^panel", g$layout$name)
  panel_cols <- unique(g$layout$l[panels])        # columns of panels
  panel_top  <- min(g$layout$t[panels])           # top row of panels
  panel_bot  <- max(g$layout$b[panels])           # bottom row of panels
  
  # add a separator at the right edge of each panel except the last
  for(i in head(panel_cols, -1)){
    g <- gtable::gtable_add_grob(
      g,
      grobs = grid::segmentsGrob(x0 = unit(1, "npc"), x1 = unit(1, "npc"),
                                 y0 = unit(0, "npc"), y1 = unit(1, "npc"),
                                 gp  = grid::gpar(col = colour, lty = linetype, lwd = size)),
      t = panel_top, b = panel_bot, l = i, r = i
    )
  }
  g
}
# Unweighted row=13, Weighted row=15, weekly count rows=17,20,23,26; Total column B
extract_weekly_fixed <- function(df, col_letters,
                                 total_col_letter = "B",
                                 uw_row = 13, wb_row = 15, count_rows = c(17,20,23,26),
                                 header_row = 11) {
  r_uw <- uw_row - header_row
  r_wb <- wb_row - header_row
  r_ct <- count_rows - header_row
  
  idx       <- xl_col_index(col_letters)
  idx_total <- xl_col_index(total_col_letter)
  
  labels <- names(df)[idx]
  getv <- function(rows, j) numify(df[rows, j, drop = TRUE])
  
  weekly_counts <- sapply(idx, function(j) sum(getv(r_ct, j), na.rm = TRUE))
  weighted_base <- sapply(idx, function(j) getv(r_wb, j))
  unweighted_n  <- sapply(idx, function(j) getv(r_uw, j))
  
  total_count   <- sum(getv(r_ct, idx_total), na.rm = TRUE)
  total_base_w  <- getv(r_wb, idx_total)
  total_base_uw <- numify(getv(r_uw, idx_total))
  total_percent <- 100 * total_count / total_base_w
  
  tibble(
    label     = labels,
    percent   = 100 * weekly_counts / weighted_base,
    n         = as.numeric(unweighted_n),
    successes = as.numeric(weekly_counts),
    total_p   = as.numeric(total_percent),
    total_n   = as.numeric(total_base_uw),
    total_s   = as.numeric(total_count)
  )
}

# Significance vs TOTAL (binom) or COMPLEMENT (Fisher), 3pp practical threshold
add_sig <- function(df, compare_to = c("total","complement"),
                    alpha = 0.05, min_diff_pp = 3) {
  compare_to <- match.arg(compare_to)
  tp <- df$total_p[1]; tn <- df$total_n[1]; ts <- df$total_s[1]
  
  out <- if (compare_to == "total") {
    df %>% rowwise() %>%
      mutate(pval = tryCatch(binom.test(successes, n, p = tp/100)$p.value, error = function(e) NA_real_),
             ref  = tp) %>%
      ungroup()
  } else {
    df %>%
      mutate(s_comp = pmax(ts - successes, 0),
             n_comp = pmax(tn - n, 0),
             p_comp = 100 * s_comp / pmax(n_comp, 1e-8)) %>%
      rowwise() %>%
      mutate(pval = tryCatch(
        fisher.test(matrix(c(successes, n - successes, s_comp, n_comp - s_comp), nrow = 2))$p.value,
        error = function(e) NA_real_
      ),
      ref = p_comp) %>%
      ungroup()
  }
  
  out %>%
    mutate(diff_pp = percent - ref,
           direction = ifelse(diff_pp > 0, "higher", "lower"),
           sig = !is.na(pval) & pval < alpha & abs(diff_pp) >= min_diff_pp)
}

# Faceted chart: subgroup labels on x-axis, single category label BELOW (strip at bottom)
plot_faceted_sig <- function(dat, title, outfile,
                             outdir = "Our changing travel/Graphs") {
  
  dat <- dat %>%
    group_by(group) %>%
    mutate(label = factor(label, levels = unique(label))) %>%
    ungroup() %>%
    mutate(fill_group = case_when(
      sig & direction == "higher" ~ "higher",
      sig & direction == "lower"  ~ "lower",
      TRUE                        ~ "ns"
    ))
  
  p <- ggplot(dat, aes(x = label, y = percent, fill = fill_group)) +
    geom_col(width = 0.75) +
    geom_text(aes(label = round(percent)), vjust = -0.35, size = 3) +
    scale_fill_manual(values = c(higher="#2C7BB6", ns="#9ECAE1", lower="#F28E2B"),
                      breaks = c("higher","ns","lower"),
                      labels = c("Higher than comparison","No significant difference","Lower than comparison"),
                      name = NULL) +
    labs(x = NULL, y = "% travelling ≥ weekly on major ‘A’/motorways", title = title) +
    scale_y_continuous(limits = c(0, 100), expand = expansion(mult = c(0, .12))) +
    facet_grid(~ group, scales = "free_x", space = "free_x", switch = "x") +
    theme_minimal(base_size = 11) +
    theme(
      panel.grid.minor = element_blank(),
      axis.text.x = element_text(angle = 45, hjust = 1),
      strip.background = element_blank(),
      strip.placement  = "outside",              # category labels BELOW the ticks
      strip.text.x.bottom = element_text(face = "bold", margin = margin(t = 6)),
      panel.spacing.x = unit(1, "lines")
    )
  
  # add one dashed vertical line between category blocks
  g <- add_facet_separators(p, colour = "grey60", linetype = "dashed", size = 0.6)
  
  if (!dir.exists(outdir)) dir.create(outdir, recursive = TRUE, showWarnings = FALSE)
  ggsave(file.path(outdir, outfile), plot = g, width = 12, height = 7, dpi = 300)
  message("Saved to: ", normalizePath(file.path(outdir, outfile), winslash = "/"))
  
  invisible(g)
}

# ------------------ MAIN ------------------
our_changing_travel_weekly_facets <- function(
    path = "Our changing travel/Our Changing Travel 0226-0227.xlsx",
    sheet_0226 = "0226", sheet_0227 = "0227",
    compare_to = c("total","complement")
){
  compare_to <- match.arg(compare_to)
  
  df26 <- read_oct(path, sheet_0226)
  df27 <- read_oct(path, sheet_0227)
  
  # Columns per your workbook:
  age_cols    <- c("C","D","E","F","G")     # 16-24, 25-34, 35-44, 45-54, 55-75
  gender_cols <- c("W","X")                 # Male, Female
  ur_cols     <- c("Y","Z")                 # Urban, Rural
  
  eth_cols    <- c("C","D")                 # Ethnic minority, White
  health_cols <- c("J","K")                 # Physical or Mental Condition, No Health Condition
  emp_cols    <- c("S","T")                 # Working FT/PT/SE, NOT FT/PT/SE
  grade_cols  <- c("Z","AA","AB","AC","AD","AE")  # AB, C1, C2, DE, ABC1, C2DE
  
  # Extract + significance (each vs the corresponding sheet total)
  age     <- extract_weekly_fixed(df26, age_cols)     |> add_sig(compare_to) |> mutate(group = "Age")
  gender  <- extract_weekly_fixed(df26, gender_cols)  |> add_sig(compare_to) |> mutate(group = "Gender")
  urban   <- extract_weekly_fixed(df26, ur_cols)      |> add_sig(compare_to) |> mutate(group = "Urban/Rural")
  
  eth     <- extract_weekly_fixed(df27, eth_cols)     |> add_sig(compare_to) |> mutate(group = "Ethnicity")
  health  <- extract_weekly_fixed(df27, health_cols)  |> add_sig(compare_to) |> mutate(group = "Health")
  emp     <- extract_weekly_fixed(df27, emp_cols)     |> add_sig(compare_to) |> mutate(group = "Employment")
  grade   <- extract_weekly_fixed(df27, grade_cols)   |> add_sig(compare_to) |> mutate(group = "Social grade")
  
  # Set A
  datA <- bind_rows(eth, age, health, gender)
  plot_faceted_sig(
    datA,
    title   = paste0("≥ weekly SRN use — Set A (comparison: ", compare_to, ")"),
    outfile = paste0("oct_weekly_setA_facets_", compare_to, ".png")
  )
  
  # Set B
  datB <- bind_rows(grade, emp, urban)
  plot_faceted_sig(
    datB,
    title   = paste0("≥ weekly SRN use — Set B (comparison: ", compare_to, ")"),
    outfile = paste0("oct_weekly_setB_facets_", compare_to, ".png")
  )
  
  invisible(list(setA = datA, setB = datB))
}

# ---- Run (pick one or both) ----
our_changing_travel_weekly_facets(compare_to = "total")
our_changing_travel_weekly_facets(compare_to = "complement")
