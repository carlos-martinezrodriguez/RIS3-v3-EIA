# =========================
# NTS — car-trip share (all slices in one chart)
# =========================
library(readxl)
library(dplyr)
library(tidyr)
library(stringr)
library(ggplot2)

# ---------- helpers ----------
numify <- function(x) as.numeric(gsub("[^0-9.\\-]", "", as.character(x)))
first_col_like <- function(df, pattern) {
  hit <- names(df)[grepl(pattern, names(df), ignore.case = TRUE)]
  if (!length(hit)) stop("Column not found: ", pattern)
  hit[1]
}
filter_year <- function(df, year_key) {
  ycol <- first_col_like(df, "^Year")
  y <- df[[ycol]]
  if (is.numeric(y)) dplyr::filter(df, .data[[ycol]] == as.numeric(year_key))
  else dplyr::filter(df, stringr::str_detect(.data[[ycol]], fixed(year_key)))
}

# car share % = (car driver + car passenger) / All modes
car_share_percent <- function(df) {
  nms <- names(df)
  allm <- nms[grepl("^All\\s*modes", nms, ignore.case = TRUE)]
  drv  <- nms[grepl("^Car\\s*or\\s*van\\s*driver", nms, ignore.case = TRUE)][1]
  pas  <- nms[grepl("^Car\\s*or\\s*van\\s*passenger", nms, ignore.case = TRUE)][1]
  if (is.na(drv) || is.na(pas)) stop("Car driver/passenger columns not found")
  car <- numify(df[[drv]]) + numify(df[[pas]])
  
  den <- if (length(allm)) numify(df[[allm[1]]]) else {
    # sum mutually exclusive mode columns if "All modes" missing
    mode_prefixes <- c("^Walk(\\b|\\s)", "^Pedal\\s*cycle", "^Car\\s*or\\s*van\\s*driver",
                       "^Car\\s*or\\s*van\\s*passenger", "^Motorcycle",
                       "^Other\\s*private\\s*transport",
                       "^Bus\\s*in\\s*London", "^Other\\s*local\\s*bus", "^Non\\-local\\s*bus",
                       "^Surface\\s*rail|^Rail(?!.*London)", "^Underground|^London\\s*Underground",
                       "^Taxi|^Minicab|^Taxi\\s*or\\s*minicab")
    mode_cols <- unique(unlist(lapply(mode_prefixes, \(p) nms[grepl(p, nms, ignore.case = TRUE)])))
    if (!length(mode_cols)) stop("No mode columns found for denominator")
    rowSums(as.data.frame(lapply(df[mode_cols], numify)), na.rm = TRUE)
  }
  100 * car / den
}

# significance vs TOTAL or COMPLEMENT (BH-adjusted), with ≥ min_diff_pp rule
add_sig <- function(df, p_total, n_total,
                    compare_to = c("total","complement"),
                    alpha = 0.05, min_diff_pp = 3, p_adjust = "BH") {
  compare_to <- match.arg(compare_to)
  out <- df
  if (compare_to == "total") {
    out <- out %>%
      mutate(
        successes = round(percent/100 * n),
        pval = vapply(seq_len(n()), function(i)
          tryCatch(binom.test(successes[i], n[i], p = p_total/100)$p.value, error = function(e) NA_real_),
          numeric(1)),
        ref = p_total
      )
  } else {
    # vs complement: derive p_comp for each row from global p_total & n_total
    out <- out %>%
      mutate(
        p_comp = (p_total*n_total - percent*n) / pmax(n_total - n, 1e-8),
        s_g = round(percent/100 * n),
        s_c = round(p_comp/100 * (n_total - n)),
        pval = vapply(seq_len(n()), function(i)
          tryCatch(fisher.test(matrix(c(s_g[i], n[i]-s_g[i], s_c[i], (n_total-n[i])-s_c[i]), nrow=2))$p.value,
                   error = function(e) NA_real_), numeric(1)),
        ref = p_comp
      )
  }
  out %>%
    mutate(
      diff_pp  = percent - ref,
      direction = ifelse(diff_pp > 0, "higher", "lower"),
      pval_adj = if (p_adjust == "none") pval else p.adjust(pval, method = p_adjust),
      sig = !is.na(pval_adj) & pval_adj < alpha & abs(diff_pp) >= min_diff_pp
    )
}

# unified plot (all slices in one chart, faceted)
plot_all <- function(dat, title, outfile, outdir = "NTS/NTS Graphs") {
  dat <- dat %>%
    mutate(fill_cat = case_when(sig & direction=="higher" ~ "higher",
                                sig & direction=="lower"  ~ "lower",
                                TRUE                      ~ "ns"),
           fill_cat = factor(fill_cat, levels = c("higher","ns","lower")),
           # keep order within each facet as provided
           label = forcats::fct_inorder(label))
  
  p <- ggplot(dat, aes(label, percent, fill = fill_cat)) +
    geom_col(width = 0.75) +
    geom_text(aes(label = round(percent)), vjust = -0.35, size = 3) +
    scale_fill_manual(
      breaks = c("higher","ns","lower"),
      values = c(higher="#2C7BB6", ns="#9ECAE1", lower="#F28E2B"),
      labels = c("Higher than comparison","No significant difference","Lower than comparison"),
      name = NULL
    ) +
    labs(x = NULL, y = "% of trips by car/van", title = title) +
    scale_y_continuous(limits = c(0, 100)) +
    theme_minimal(base_size = 11) +
    theme(panel.grid.minor = element_blank(),
          axis.text.x = element_text(angle = 45, hjust = 1)) +
    facet_wrap(~category, scales = "free_x", nrow = 2)
  
  if (!dir.exists(outdir)) dir.create(outdir, recursive = TRUE, showWarnings = FALSE)
  ggsave(file.path(outdir, outfile), plot = p, width = 12, height = 7, dpi = 300)
  message("Saved to: ", normalizePath(file.path(outdir, outfile), winslash = "/"))
  p
}

# ---------- loaders ----------
# 1) Gender & Age
load_gender_age_2024 <- function(path = "NTS/Surveys/NTS trips by sex, age and main mode.xlsx",
                                 sheet = "NTS0601a_trips",
                                 year_key = "2024") {
  df <- read_xlsx(path, sheet = sheet, skip = 5) |> filter_year(year_key)
  
  sex_col  <- first_col_like(df, "^Sex")
  age_col  <- first_col_like(df, "^Age")
  base_col <- first_col_like(df, "^Unweighted\\s*sample\\s*size:\\s*individuals")
  
  df$percent <- car_share_percent(df)
  df$n       <- numify(df[[base_col]])
  
  overall <- df %>%
    filter(.data[[sex_col]] == "All people", .data[[age_col]] == "All ages") %>%
    summarise(p_total = dplyr::first(percent), n_total = dplyr::first(n), .groups="drop")
  
  gender <- df %>%
    filter(.data[[age_col]] == "All ages", .data[[sex_col]] %in% c("Males","Females")) %>%
    transmute(category = "Gender", label = .data[[sex_col]], percent, n)
  
  age <- df %>%
    filter(.data[[sex_col]] == "All people", .data[[age_col]] != "All ages") %>%
    transmute(category = "Age", label = .data[[age_col]], percent, n)
  
  list(gender = gender, age = age,
       p_total_nat = overall$p_total[1], n_total_nat = overall$n_total[1])
}

# 2) Ethnicity (aggregate regions to national)
load_ethnicity_2024 <- function(path = "NTS/Surveys/NTS trips and distance travelled by mode, ethnic group.xlsx",
                                sheet = "NTS9917a_trips_region",
                                year_key = "2024") {
  df <- read_xlsx(path, sheet = sheet, skip = 5) |> filter_year(year_key)
  
  eth_col  <- first_col_like(df, "^Ethnic\\s*group")
  base_col <- first_col_like(df, "^Unweighted\\s*sample\\s*size:\\s*individuals")
  
  df$percent <- car_share_percent(df)
  df$n       <- numify(df[[base_col]])
  
  eth_nat <- df %>%
    filter(.data[[eth_col]] %in% c("White","Ethnic minorities")) %>%
    group_by(.data[[eth_col]]) %>%
    summarise(percent = sum(percent * n, na.rm = TRUE)/sum(n, na.rm = TRUE),
              n = sum(n, na.rm = TRUE), .groups = "drop") %>%
    transmute(category = "Ethnicity", label = .data[[eth_col]], percent, n)
  
  list(ethnicity = eth_nat)
}

# 3) Disability (use All aged 16+ if present)
load_disability_2024 <- function(path = "NTS/Surveys/NTS trips and distance travelled by mode, disability status and age.xlsx",
                                 sheet = "DIS0402a_trips",
                                 year_key = "2024") {
  df <- read_xlsx(path, sheet = sheet, skip = 5) |> filter_year(year_key)
  
  age_col  <- first_col_like(df, "^Age\\s*group")
  dis_col  <- first_col_like(df, "^Disability\\s*status")
  base_col <- first_col_like(df, "^Unweighted\\s*sample\\s*size:\\s*individuals")
  
  df$percent <- car_share_percent(df)
  df$n       <- numify(df[[base_col]])
  
  df_allage <- df %>% filter(.data[[age_col]] %in% c("All aged 16 and over","All ages 16 and over"))
  if (nrow(df_allage) == 0) df_allage <- df
  
  dis <- df_allage %>%
    filter(.data[[dis_col]] %in% c("Disabled","Not disabled")) %>%
    group_by(.data[[dis_col]]) %>%
    summarise(percent = sum(percent * n, na.rm = TRUE)/sum(n, na.rm = TRUE),
              n = sum(n, na.rm = TRUE), .groups = "drop") %>%
    transmute(category = "Disability", label = .data[[dis_col]], percent, n)
  
  list(disability = dis)
}

# ---------- build, test, plot ----------
# One-panel chart
plot_all_carshare_2024_side_by_side <- function(compare_to = c("total","complement"),
                                                year_key = "2024",
                                                outdir = "NTS/NTS Graphs") {
  compare_to <- match.arg(compare_to)
  
  # load slices 
  ga  <- load_gender_age_2024(year_key = year_key)
  eth <- load_ethnicity_2024(year_key = year_key)
  dis <- load_disability_2024(year_key = year_key)
  
  combined <- dplyr::bind_rows(ga$gender, ga$age, eth$ethnicity, dis$disability)
  
  # overall national p and n from All people + All ages (gender/age table)
  p_total <- ga$p_total_nat
  n_total <- ga$n_total_nat
  
  dat <- add_sig(combined, p_total = p_total, n_total = n_total,
                 compare_to = compare_to, min_diff_pp = 3)
  
  # order blocks on the x-axis (change order here if you like)
  cat_order <- c("Age","Disability","Ethnicity","Gender")
  dat <- dat %>%
    dplyr::mutate(category = factor(category, levels = cat_order)) %>%
    dplyr::group_by(category) %>%
    dplyr::mutate(label = factor(label, levels = unique(label))) %>%
    dplyr::ungroup() %>%
    dplyr::arrange(category, label) %>%
    dplyr::mutate(
      xlab = factor(paste0(as.character(label), "\n", as.character(category)),
                    levels = paste0(as.character(label), "\n", as.character(category))),
      fill_cat = dplyr::case_when(
        sig & direction == "higher" ~ "higher",
        sig & direction == "lower"  ~ "lower",
        TRUE                        ~ "ns"
      ),
      fill_cat = factor(fill_cat, levels = c("higher","ns","lower"))
    )
  
  # positions for faint separators between category blocks
  blocks <- dat |>
    dplyr::count(category, name = "n") |>
    dplyr::mutate(end = cumsum(n), start = end - n + 1L)
  cutpos <- (blocks$end[-nrow(blocks)] + 0.5)
  
  p <- ggplot2::ggplot(dat, ggplot2::aes(x = xlab, y = percent, fill = fill_cat)) +
    ggplot2::geom_col(width = 0.75) +
    ggplot2::geom_text(ggplot2::aes(label = round(percent)), vjust = -0.35, size = 3) +
    ggplot2::geom_vline(xintercept = cutpos, linetype = "dashed", alpha = 0.3) +
    ggplot2::scale_fill_manual(
      breaks = c("higher","ns","lower"),
      values = c(higher = "#2C7BB6", ns = "#9ECAE1", lower = "#F28E2B"),
      labels = c("Higher than comparison", "No significant difference", "Lower than comparison"),
      name = NULL
    ) +
    ggplot2::labs(
      x = NULL, y = "% of trips by car/van",
      title = paste0("Car trip share — Gender, Age, Ethnicity, Disability (", year_key, ")\n",
                     "Comparison: ", compare_to)
    ) +
    ggplot2::scale_y_continuous(limits = c(0, 100)) +
    ggplot2::theme_minimal(base_size = 11) +
    ggplot2::theme(
      panel.grid.minor = ggplot2::element_blank(),
      axis.text.x = ggplot2::element_text(angle = 45, hjust = 1)
    )
  
  if (!dir.exists(outdir)) dir.create(outdir, recursive = TRUE, showWarnings = FALSE)
  outfile <- file.path(outdir, paste0("nts_carshare_all_side_by_side_", year_key, "_", compare_to, ".png"))
  ggplot2::ggsave(outfile, plot = p, width = 12, height = 6.5, dpi = 300)
  message("Saved to: ", normalizePath(outfile, winslash = "/"))
  p
}


# 
# Plotting.
plot_all_carshare_2024_side_by_side(compare_to = "total")
plot_all_carshare_2024_side_by_side(compare_to = "complement")
