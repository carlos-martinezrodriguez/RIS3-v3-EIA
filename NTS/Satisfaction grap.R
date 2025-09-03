library(readxl)
library(dplyr)
library(tidyr)
library(ggplot2)
library(stringr)

nts_satisfaction_plot <- function(
    path = "NTS/Surveys/NTS Satisfaction with SRN.xlsx",
    sheet = "NTS0802d_major_roads",
    outdir = "NTS/NTS Graphs",
    outfile = "nts_satisfaction_srn.png"
){
  # read: skip 4 header rows with notes so row 5 becomes header
  df <- read_xlsx(path, sheet = sheet, skip = 4)
  
  # pick response columns; drop "Satisfaction" label, "All users", and sample size
  drop_cols <- grep("All users|Unweighted|Satisfaction", names(df), value = TRUE)
  measure_cols <- setdiff(names(df), c("Year", drop_cols))
  
  long <- df %>%
    select(Year, all_of(measure_cols)) %>%
    pivot_longer(-Year, names_to = "response", values_to = "percent") %>%
    mutate(
      percent = as.numeric(percent),
      response = str_replace_all(response, "\\s*\\(%\\)", "") |> str_squish(),
      # set a sensible legend/order
      response = factor(response,
                        levels = c("Very satisfied", "Fairly satisfied",
                                   "Neither satisfied nor dissatisfied",
                                   "Fairly dissatisfied", "Very dissatisfied"))
    ) %>%
    arrange(response, Year)
  
  p <- ggplot(long, aes(x = Year, y = percent, color = response)) +
    geom_line(linewidth = 1) +
    geom_point(size = 2) +
    scale_y_continuous(limits = c(0, 60)) +
    labs(x = NULL, y = "% of respondents",
         title = "Satisfaction with provision of major roads (NTS)",
         subtitle = "England, 2016 onwards",
         color = NULL) +
    theme_minimal(base_size = 11) +
    theme(panel.grid.minor = element_blank())
  
  if (!dir.exists(outdir)) dir.create(outdir, recursive = TRUE)
  ggsave(file.path(outdir, outfile), plot = p, width = 8, height = 5, dpi = 300)
  
  message("Saved: ", file.path(outdir, outfile))
  invisible(p)
}

# Run it
nts_satisfaction_plot()