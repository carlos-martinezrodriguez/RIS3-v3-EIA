library(readxl)
library(dplyr)
library(tidyr)
library(ggplot2)
library(stringr)

plot_ethnicity_car_access <- function(
    path = "NTS/Surveys/NTS personal car access and trip rates, by ethnic group.xlsx",
    sheet = "NTS0707a_personal_car_access", 
    year_filter = "2023 to 2024",
    outdir = "NTS/NTS Graphs",
    outfile = "nts_ethnicity_car_access_2024.png"
){
  df <- read_xlsx(path, sheet = sheet, skip = 4)
  
  # filter year 2024
  df_2024 <- df %>% filter(Year == year_filter)
  
  # pick "All adults with a car or van (%)" column (col G in screenshot)
  df_sel <- df_2024 %>%
    select(`Ethnic group`, percent = `All adults in households with a car or van (%)`,
           n = `Unweighted sample size: individuals aged 17 and over (number) [note 2]`)
  
  # overall row
  overall <- df_sel %>% filter(str_detect(`Ethnic group`, "All ethnic")) %>%
    summarise(p_total = percent, n_total = n)
  
  # significance: binomial test each group vs overall
  results <- df_sel %>%
    filter(!str_detect(`Ethnic group`, "All ethnic")) %>%
    rowwise() %>%
    mutate(
      s_g = round(percent/100 * n),
      pval = tryCatch(binom.test(s_g, n, p = overall$p_total/100)$p.value, error = function(e) NA),
      diff_pp = percent - overall$p_total,
      direction = ifelse(percent > overall$p_total, "higher", "lower"),
      sig = !is.na(pval) & pval < 0.05 & abs(diff_pp) >= 3   # require ≥3pp gap
    ) %>% ungroup()
  
  # plot
  p <- ggplot(results, aes(x = `Ethnic group`, y = percent,
                           fill = case_when(
                             sig & direction == "higher" ~ "higher",
                             sig & direction == "lower"  ~ "lower",
                             TRUE                        ~ "ns"
                           ))) +
    geom_col(width = 0.7) +
    geom_text(aes(label = round(percent)), vjust = -0.3, size = 3) +
    scale_fill_manual(
      breaks = c("higher","ns","lower"),
      values = c(higher = "#2C7BB6", ns = "#9ECAE1", lower = "#F28E2B"),
      labels = c(
        higher = "Higher than average",
        ns     = "No significant difference",
        lower  = "Lower than average"
      ),
      name = NULL) +
    labs(x = NULL, y = "% adults with car/van access",
         title = "Car/van access by ethnic group (2024)",
         caption = "Comparison vs All ethnic groups; binomial test p<0.05, ≥3pp difference") +
    scale_y_continuous(limits = c(0,100)) +
    theme_minimal(base_size = 11) +
    theme(axis.text.x = element_text(angle = 45, hjust = 1),
          panel.grid.minor = element_blank())
  
  if (!dir.exists(outdir)) dir.create(outdir, recursive = TRUE)
  ggsave(file.path(outdir, outfile), plot = p, width = 8, height = 5, dpi = 300)
  
  message("Saved: ", file.path(outdir, outfile))
  invisible(p)
}

# Run it
plot_ethnicity_car_access()
