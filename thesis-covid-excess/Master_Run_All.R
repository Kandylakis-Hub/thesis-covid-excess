# Master_Run_All.R
# ============================================================
# ΕΝΙΑΙΟ SCRIPT: Περιγραφικά & Γραφικά + Μοντέλα (0–8, 0–26)
# ΣΗΜΑΝΤΙΚΟ: ΔΕΝ αλλάζουμε τίποτα σε plots/λεζάντες/lags/NW-lags.
# Κρατάμε τα ίδια paths που χρησιμοποιείς.
# ============================================================

suppressPackageStartupMessages({
  library(readxl); library(dplyr); library(tidyr); library(stringr)
  library(tibble); library(lmtest); library(sandwich); library(broom)
  library(writexl); library(ggplot2); library(purrr); library(scales)
})

# ---- Paths ----
###  Σημείωση για το αρχείο δεδομένων
#Το αρχείο `Final Data.xlsx` πρέπει να βρίσκεται **στον ίδιο φάκελο** με το script `Master_Run_All.R`.  
#Αν χρησιμοποιείτε διαφορετική διαδρομή, αλλάξτε τις μεταβλητές `data_path` και `out_dir` 
#στις πρώτες γραμμές του script.
root_dir <- getwd()  # ο φάκελος όπου βρίσκεται το script
data_path <- file.path(root_dir, "Final Data.xlsx")
out_dir   <- root_dir
plot_dir  <- file.path(out_dir, "plots")

# Δημιουργία φακέλων αν δεν υπάρχουν
if (!dir.exists(out_dir))  dir.create(out_dir, recursive = TRUE)
if (!dir.exists(plot_dir)) dir.create(plot_dir, recursive = TRUE)

message("📂 Χρησιμοποιείται root directory: ", normalizePath(root_dir))
message("📄 Αναμένεται αρχείο δεδομένων: ", data_path)

# ---- Global theme: ΛΕΥΚΟ ----
theme_set(
  theme_minimal(base_size = 12) +
    theme(
      plot.background  = element_rect(fill = "white", colour = NA),
      panel.background = element_rect(fill = "white", colour = NA),
      legend.background= element_rect(fill = "white", colour = NA),
      legend.key       = element_rect(fill = "white", colour = NA),
      legend.position  = "bottom",
      plot.caption.position = "plot",
      axis.text.x = element_text(angle = 0, hjust = 0.5, vjust = 0.5)
    )
)

# ---- Σταθερή παλέτα ανά χώρα / code (όπως είχες) ----
country_palette <- c(
  "Greece"      = "#66B2FF", "Ελλάδα"   = "#66B2FF", "GRC" = "#66B2FF",
  "Belgium"     = "#E41A1C", "BEL"      = "#E41A1C",
  "Netherlands" = "#FF8C00", "NLD"      = "#FF8C00",
  "Bulgaria"    = "#2CA02C", "BGR"      = "#2CA02C",
  "Iceland"     = "#003366", "ISL"      = "#003366"
)
country_palette_code <- c(GRC="#66B2FF", BEL="#E41A1C", NLD="#FF8C00", BGR="#2CA02C", ISL="#003366")
country_color      <- function(cc) if (cc %in% names(country_palette)) country_palette[[cc]] else "#333333"
get_country_color  <- function(cc) if (!is.null(country_palette_code[[cc]])) country_palette_code[[cc]] else "#555555"

# ============================================================
#               1) ΠΕΡΙΓΡΑΦΙΚΑ & ΓΡΑΦΙΚΑ
# ============================================================

df <- read_xlsx(data_path)
names(df) <- trimws(names(df))

country_col <- dplyr::case_when(
  "CountryName" %in% names(df) ~ "CountryName",
  "Country" %in% names(df) ~ "Country",
  "Country_Name" %in% names(df) ~ "Country_Name",
  "Countryname" %in% names(df) ~ "Countryname",
  "CountryCode" %in% names(df) ~ "CountryCode",
  "ISO3" %in% names(df) ~ "ISO3",
  TRUE ~ NA_character_
)
if (is.na(country_col)) stop("Δεν βρέθηκε στήλη χώρας.")

if ((!("Year" %in% names(df)) || !("Week" %in% names(df))) && ("Week_Index" %in% names(df))) {
  parts <- str_split_fixed(as.character(df$Week_Index), "_", 2)
  if (!("Year" %in% names(df))) df$Year <- as.integer(parts[,1])
  if (!("Week" %in% names(df)) && ncol(parts) == 2) df$Week <- as.integer(parts[,2])
}
stopifnot(all(c("Year","Week") %in% names(df)))

df <- df %>% mutate(Half = if_else(Week <= 26, "H1", "H2"))

metrics <- c("Total_Deaths","Expected_Deaths","Excess_Deaths",
             "Excess_percent","Deaths_per_100k","Excess_per_100k")

policy_map <- list(
  StringencyIndex_Average          = c("StringencyIndex_Average","StringencyIndex","StringencyIndex_Mean","Stringency_Index"),
  GovernmentResponseIndex_Average  = c("GovernmentResponseIndex_Average","GovernmentResponseIndex","GovResponseIndex"),
  ContainmentHealthIndex_Average   = c("ContainmentHealthIndex_Average","ContainmentHealthIndex","Containment_Index"),
  EconomicSupportIndex             = c("EconomicSupportIndex","EconomicSupportIndex_Average","Economic_Support_Index")
)
resolve_col <- function(cands) { hit <- cands[cands %in% names(df)]; if (length(hit)==0) NA_character_ else hit[1] }
policy_cols <- purrr::imap_chr(policy_map, ~ resolve_col(.x))
for (nm in names(policy_cols)) if (is.na(policy_cols[[nm]])) { df[[nm]] <- NA_real_; policy_cols[[nm]] <- nm }

to_numeric <- unique(c(metrics, unname(policy_cols)))
for (cl in to_numeric) if (cl %in% names(df)) df[[cl]] <- suppressWarnings(as.numeric(df[[cl]]))

# ---- Πίνακες (σε πολλά sheets) ----
agg_funs <- list(mean = ~mean(.x, na.rm = TRUE),
                 median = ~median(.x, na.rm = TRUE),
                 sd = ~sd(.x, na.rm = TRUE))

desc_by_country_year <- df %>%
  group_by(.data[[country_col]], Year) %>%
  summarise(across(all_of(metrics), agg_funs, .names = "{.col}_{.fn}"), .groups = "drop") %>%
  arrange(.data[[country_col]], Year)

policy_by_year <- df %>%
  group_by(.data[[country_col]], Year) %>%
  summarise(across(all_of(unname(policy_cols)),
                   list(mean = ~mean(.x, na.rm = TRUE), sd = ~sd(.x, na.rm = TRUE)),
                   .names = "{.col}_{.fn}"),
            .groups = "drop") %>%
  arrange(.data[[country_col]], Year)

policy_by_year_half <- df %>%
  group_by(.data[[country_col]], Year, Half) %>%
  summarise(across(all_of(unname(policy_cols)),
                   list(mean = ~mean(.x, na.rm = TRUE), sd = ~sd(.x, na.rm = TRUE)),
                   .names = "{.col}_{.fn}"),
            .groups = "drop") %>%
  arrange(.data[[country_col]], Year, Half)

write_xlsx(
  list(
    Desc_by_Country_Year = desc_by_country_year,
    Policy_by_Year       = policy_by_year,
    Policy_by_Year_H1H2  = policy_by_year_half
  ),
  path = file.path(out_dir, "Descriptive_Summary.xlsx")
)

# ---- Προετοιμασία για γραφήματα ----
df <- df %>% arrange(.data[[country_col]], Year, Week) %>%
  mutate(
    Week_Label = if ("Week_Index" %in% names(df)) as.character(Week_Index) else sprintf("%d_%02d", Year, Week),
    seq = ave(Week, .data[[country_col]], FUN = seq_along)
  )

countries <- df %>% distinct(.data[[country_col]]) %>% pull(1) %>% as.character()
str_col <- policy_cols[["StringencyIndex_Average"]]

# ---- 1) Time series: Total vs Expected ----
plot_total_expected <- function(d, cc) {
  col_cc <- country_color(cc)
  ggplot(d, aes(x = seq)) +
    geom_line(aes(y = Total_Deaths, colour = "Σύνολο Θανάτων"), linewidth = 0.9, show.legend = TRUE) +
    geom_line(aes(y = Expected_Deaths, colour = "Αναμενόμενοι Θάνατοι"), linewidth = 0.9, colour = "black", show.legend = TRUE) +
    scale_color_manual(values = c("Σύνολο Θανάτων" = col_cc, "Αναμενόμενοι Θάνατοι" = col_cc)) +
    scale_x_continuous(breaks = pretty_breaks(n = 12),
                       labels = function(ix){ lab <- d$Week_Label[match(ix, d$seq)]; ifelse(is.na(lab),"",lab) }) +
    labs(title = paste0(cc, " — Σύνολο vs Αναμενόμενοι (εβδομαδιαία)"),
         x = "Εβδομάδες (σειριακά)", y = "Αριθμός θανάτων",
         caption = "Συνεχής γραμμή: Σύνολο θανάτων · Διακεκομμένη: Αναμενόμενοι.") +
    guides(colour = guide_legend(title = NULL))
}

# ---- 2) Excess_per_100k με overlay Stringency ----
plot_excess_with_stringency <- function(d, cc) {
  if (!("Excess_per_100k" %in% names(d)) || !(str_col %in% names(d))) return(invisible(NULL))
  col_cc <- country_color(cc)
  rng_y <- range(d$Excess_per_100k, na.rm = TRUE); rng_s <- range(d[[str_col]], na.rm = TRUE)
  if (!all(is.finite(rng_y)) || !all(is.finite(rng_s)) || diff(rng_s) == 0) return(invisible(NULL))
  scale_factor <- diff(rng_y) / diff(rng_s)
  
  ggplot(d, aes(x = seq)) +
    geom_line(aes(y = Excess_per_100k, colour = "Excess per 100k"), linewidth = 0.9) +
    geom_line(aes(y = !!sym(str_col) * scale_factor, colour = "Stringency (δεξιός άξονας)"),
              linewidth = 0.9, colour = "black") +
    scale_color_manual(values = c("Excess per 100k" = col_cc, "Stringency (δεξιός άξονας)" = "grey35")) +
    scale_y_continuous(name = "Excess per 100k",
                       sec.axis = sec_axis(~ . / scale_factor, name = "Stringency Index")) +
    scale_x_continuous(breaks = pretty_breaks(n = 12),
                       labels = function(ix){ lab <- d$Week_Label[match(ix, d$seq)]; ifelse(is.na(lab),"",lab) }) +
    labs(title = paste0(cc, " — Excess per 100k & Stringency (overlay)"),
         x = "Εβδομάδες (σειριακά)",
         caption = "Συνεχής: Excess per 100k (αριστερός άξονας) · Διακεκομμένη γκρι: Stringency (δεξιός άξονας).") +
    guides(colour = guide_legend(title = NULL))
}

# ---- 3) Boxplot Excess_per_100k ανά έτος ----
plot_box_excess_by_year <- function(d, cc) {
  col_cc <- country_color(cc)
  ggplot(d, aes(x = factor(Year), y = Excess_per_100k)) +
    geom_boxplot(outlier.alpha = 0.35, colour = col_cc, fill = alpha(col_cc, 0.12)) +
    stat_summary(fun = "mean", geom = "point", shape = 21, size = 2, fill = "white", colour = col_cc) +
    labs(title = paste0(cc, " — Boxplot Excess per 100k ανά Έτος"),
         x = "Έτος", y = "Excess per 100k",
         caption = "Το κουτί δείχνει IQR (Q1–Q3), η γραμμή τη διάμεσο, η τελεία τον μέσο όρο.")
}

# ---- 3b) Density Excess_per_100k ανά έτος ----
year_palette_from_base <- function(base_col, years) {
  pal_fn <- seq_gradient_pal("white", base_col, "Lab")
  pal_fn(seq(0.4, 1, length.out = length(years)))
}
plot_density_excess_by_year <- function(d, cc) {
  col_cc <- country_color(cc)
  yrs <- sort(unique(d$Year))
  vals <- year_palette_from_base(col_cc, yrs)
  ggplot(d, aes(x = Excess_per_100k,
                colour = factor(Year), fill = factor(Year))) +
    geom_density(alpha = 0.35) +
    scale_color_manual(values = setNames(vals, yrs), guide = guide_legend(title = "Έτος")) +
    scale_fill_manual(values = setNames(vals, yrs), guide = guide_legend(title = "Έτος")) +
    labs(title = paste0(cc, " — Density Excess per 100k ανά Έτος"),
         x = "Excess per 100k", y = "Density",
         caption = "Κάθε απόχρωση του χρώματος χώρας αντιστοιχεί σε διαφορετικό έτος (βλ. υπόμνημα κάτω).")
}

# ---- 4) Scatter Excess_per_100k vs Stringency ----
plot_scatter_excess_vs_stringency <- function(d, cc) {
  if (!("Excess_per_100k" %in% names(d)) || !(str_col %in% names(d))) return(invisible(NULL))
  d2 <- d %>% filter(!is.na(Excess_per_100k), !is.na(.data[[str_col]]))
  if (nrow(d2) == 0) return(invisible(NULL))
  col_cc <- country_color(cc)
  
  ggplot(d2, aes(x = .data[[str_col]], y = Excess_per_100k)) +
    geom_point(alpha = 0.7, colour = col_cc) +
    geom_smooth(method = "lm", formula = y ~ poly(x, 2, raw = TRUE),
                se = TRUE, colour = col_cc, fill = alpha(col_cc, 0.18)) +
    labs(title = paste0(cc, " — Excess per 100k vs Stringency"),
         x = "Stringency Index", y = "Excess per 100k",
         caption = "Η καμπύλη είναι πολυωνυμική προσαρμογή 2ου βαθμού (LM) με 95% ΔΕ.")
}

save_and_show <- function(p, path, w = 10, h = 5) {
  if (inherits(p, "ggplot")) {
    print(p)
    ggsave(filename = path, plot = p, width = w, height = h, dpi = 150, bg = "white")
  }
}

# ---- Τρέχουμε για κάθε χώρα ----
walk(countries, function(cc) {
  d_cc <- df %>% filter(.data[[country_col]] == cc) %>% arrange(Year, Week)
  
  p1 <- plot_total_expected(d_cc, cc)
  save_and_show(p1, file.path(plot_dir, paste0(cc, "_time_Total_vs_Expected.png")))
  
  p2 <- plot_excess_with_stringency(d_cc, cc)
  save_and_show(p2, file.path(plot_dir, paste0(cc, "_time_Excess100k_with_Stringency.png")))
  
  p3 <- plot_box_excess_by_year(d_cc, cc)
  save_and_show(p3, file.path(plot_dir, paste0(cc, "_boxplot_Excess100k_by_Year.png")))
  
  p3b <- plot_density_excess_by_year(d_cc, cc)
  save_and_show(p3b, file.path(plot_dir, paste0(cc, "_density_Excess100k_by_Year.png")))
  
  p4 <- plot_scatter_excess_vs_stringency(d_cc, cc)
  save_and_show(p4, file.path(plot_dir, paste0(cc, "_scatter_Excess100k_vs_Stringency.png")))
})

# ---- Ευρετήριο PNG ----
write.csv(data.frame(file = list.files(plot_dir, pattern = "\\.png$", full.names = FALSE)),
          file.path(out_dir, "plot_file_index.csv"), row.names = FALSE)

message("\n✅ Περιγραφικά έτοιμα:\n- ", file.path(out_dir, "Descriptive_Summary.xlsx"),
        "\n- Φάκελος γραφημάτων: ", normalizePath(plot_dir), "\n")

# ============================================================
#             2) Model_Stringency_Only (0..8, NW=4)
# ============================================================

theme_set(
  theme_minimal(base_size = 12) +
    theme(
      plot.background  = element_rect(fill = "white", colour = NA),
      panel.background = element_rect(fill = "white", colour = NA),
      legend.position  = "bottom",
      plot.caption.position = "plot",
      axis.text.x = element_text(angle = 0, hjust = 0.5, vjust = 0.5)
    )
)

out_xlsx_explor_08  <- file.path(out_dir, "PerCountry_SI_LagSweep.xlsx")
out_xlsx_primary_08 <- file.path(out_dir, "PerCountry_Primary_2to4.xlsx")
plot_dir_08  <- file.path(out_dir, "per_country_plots")
lags_explor  <- 0:8
start_year <- 2020; start_week <- 1
end_year   <- 2022; end_week   <- 52
nw_lag     <- 4   # Newey–West window (≈1 μήνας)
dir.create(plot_dir_08, showWarnings = FALSE, recursive = TRUE)

parse_year_week <- function(week_id_chr) {
  tibble(
    Week_Index = week_id_chr,
    Year = suppressWarnings(as.integer(stringr::str_sub(week_id_chr, 1, 4))),
    Week = suppressWarnings(as.integer(stringr::str_sub(week_id_chr, 6, 7)))
  )
}
nw_row <- function(mod, term, lag_nw = 4) {
  V  <- sandwich::NeweyWest(mod, lag = lag_nw, prewhite = FALSE, adjust = TRUE)
  ct <- lmtest::coeftest(mod, vcov. = V)
  if (!term %in% rownames(ct)) {
    return(tibble(term = term, estimate = NA_real_, std.error = NA_real_,
                  statistic = NA_real_, p.value = NA_real_))
  }
  tibble(
    term = term,
    estimate = unname(ct[term,1]),
    std.error = unname(ct[term,2]),
    statistic = unname(ct[term,3]),
    p.value = unname(ct[term,4])
  )
}

required_cols <- c(
  "CountryName","CountryCode","Year","Week","Week_Index",
  "Total_Deaths","Expected_Deaths","Excess_Deaths","Excess_percent",
  "Population","Deaths_per_100k","Excess_per_100k",
  "StringencyIndex_Average","GovernmentResponseIndex_Average",
  "ContainmentHealthIndex_Average","EconomicSupportIndex",
  "ConfirmedCases_weekly","ConfirmedDeaths_weekly",
  "COVID_Cases_per_100k","COVID_Deaths_per_100k"
)

df_raw <- read_xlsx(data_path)
miss <- setdiff(required_cols, names(df_raw))
if (length(miss) > 0) stop("Λείπουν στήλες από το Excel: ", paste(miss, collapse = ", "))

dfm <- df_raw %>%
  mutate(
    Week_Index = as.character(Week_Index),
    Year = suppressWarnings(as.integer(Year)),
    Week = suppressWarnings(as.integer(Week))
  )

if (any(is.na(dfm$Year)) || any(is.na(dfm$Week))) {
  dfm <- dfm %>% select(-Year, -Week) %>%
    left_join(parse_year_week(dfm$Week_Index), by = "Week_Index")
}

dfm <- dfm %>%
  filter(
    (Year >  start_year | (Year == start_year & Week >= start_week)) &
      (Year <  end_year   | (Year == end_year   & Week <= end_week))
  ) %>%
  group_by(CountryCode, Week_Index) %>% slice(1L) %>% ungroup()

countries_08 <- sort(unique(dfm$CountryCode))

overview_list <- list()
sheets_explor <- list()

for (cc in countries_08) {
  d_cc <- dfm %>% filter(CountryCode == cc) %>%
    arrange(Year, Week) %>%
    mutate(SI = StringencyIndex_Average)
  
  lag_rows <- list()
  col_cc <- get_country_color(cc)
  
  for (k in lags_explor) {
    d_k <- d_cc %>%
      mutate(SI10_lag = dplyr::lag(SI, k) / 10) %>%        # Stringency/10 με lag k
      filter(!is.na(Excess_per_100k), !is.na(SI10_lag))
    
    if (nrow(d_k) < 20) {
      lag_rows[[as.character(k)]] <- tibble(
        CountryCode = cc, lag = k,
        beta = NA_real_, se = NA_real_, t = NA_real_, p = NA_real_,
        conf_low = NA_real_, conf_high = NA_real_,
        N = nrow(d_k), R2 = NA_real_, AdjR2 = NA_real_
      )
      next
    }
    
    mod <- lm(Excess_per_100k ~ SI10_lag + factor(Week) + factor(Year), data = d_k)
    
    rob <- nw_row(mod, "SI10_lag", lag_nw = nw_lag)
    z <- qnorm(0.975)
    conf_low  <- rob$estimate - z * rob$std.error
    conf_high <- rob$estimate + z * rob$std.error
    gl <- broom::glance(mod)
    
    lag_rows[[as.character(k)]] <- tibble(
      CountryCode = cc, lag = k,
      beta = rob$estimate, se = rob$std.error, t = rob$statistic, p = rob$p.value,
      conf_low = conf_low, conf_high = conf_high,
      N = nobs(mod), R2 = gl$r.squared, AdjR2 = gl$adj.r.squared
    )
  }
  
  lag_tbl <- bind_rows(lag_rows) %>% arrange(lag) %>%
    mutate(p_fdr = p.adjust(p, method = "BH"))
  
  overview_list[[cc]] <- lag_tbl
  sheets_explor[[paste0("Lag_", cc)]] <- lag_tbl
  
  # Exploratory plot (0–8)  (μόνο beta/CI)
  p <- ggplot(lag_tbl, aes(lag, beta)) +
    geom_hline(yintercept = 0, linetype = "dashed") +
    geom_point(color = col_cc) +
    geom_errorbar(aes(ymin = conf_low, ymax = conf_high), width = 0.15, color = col_cc) +
    scale_x_continuous(breaks = lags_explor) +
    labs(
      title = paste0(cc, ": Stringency/10 (lag) → Excess_per_100k"),
      subtitle = "FE(Week) + FE(Year), Newey–West SEs — Exploratory 0–8",
      x = "Lag (εβδομάδες)",
      y = "Beta (μονάδες Excess/100k ανά +10 μονάδες Stringency)"
    )
  ggsave(file.path(plot_dir_08, paste0("Lag_Beta_CI_", cc, ".png")),
         plot = p, width = 8, height = 4.8, dpi = 300)
}

overview <- bind_rows(overview_list) %>%
  select(CountryCode, lag, beta, se, t, p, p_fdr, conf_low, conf_high, N, R2, AdjR2)

write_xlsx(c(list("Overview" = overview), sheets_explor), path = out_xlsx_explor_08)

# PRIMARY (2–4)
primary_rows <- list()
for (cc in countries_08) {
  d_cc <- dfm %>% filter(CountryCode == cc) %>%
    arrange(Year, Week) %>%
    mutate(SI = StringencyIndex_Average)
  
  d_p <- d_cc %>%
    mutate(
      SI_lag2 = dplyr::lag(SI, 2),
      SI_lag3 = dplyr::lag(SI, 3),
      SI_lag4 = dplyr::lag(SI, 4),
      SI10_2to4 = rowMeans(cbind(SI_lag2, SI_lag3, SI_lag4), na.rm = FALSE) / 10
    ) %>%
    filter(!is.na(Excess_per_100k), !is.na(SI10_2to4))
  
  if (nrow(d_p) < 20) {
    primary_rows[[cc]] <- tibble(
      CountryCode = cc, beta = NA_real_, se = NA_real_, t = NA_real_, p = NA_real_,
      conf_low = NA_real_, conf_high = NA_real_, N = nrow(d_p),
      R2 = NA_real_, AdjR2 = NA_real_
    )
    next
  }
  
  mod_p <- lm(Excess_per_100k ~ SI10_2to4 + factor(Week) + factor(Year), data = d_p)
  rob_p <- nw_row(mod_p, "SI10_2to4", lag_nw = nw_lag)
  z <- qnorm(0.975)
  conf_low  <- rob_p$estimate - z * rob_p$std.error
  conf_high <- rob_p$estimate + z * rob_p$std.error
  gl <- broom::glance(mod_p)
  
  primary_rows[[cc]] <- tibble(
    CountryCode = cc,
    beta = rob_p$estimate, se = rob_p$std.error, t = rob_p$statistic, p = rob_p$p.value,
    conf_low = conf_low, conf_high = conf_high,
    N = nobs(mod_p), R2 = gl$r.squared, AdjR2 = gl$adj.r.squared
  )
}
primary_tbl <- bind_rows(primary_rows) %>% arrange(CountryCode)
write_xlsx(list("Primary_2to4" = primary_tbl), path = out_xlsx_primary_08)

message("✅ Έτοιμο (0–8, NW=4).",
        "\n- Exploratory: ", out_xlsx_explor_08,
        "\n- Primary (2–4): ", out_xlsx_primary_08,
        "\n- Plots: ", plot_dir_08, "\n")

# ============================================================
#     3) Results_StringencyOnly26W (0..26, NW=8)
# ============================================================

theme_set(
  theme_minimal(base_size = 12) +
    theme(
      plot.background  = element_rect(fill = "white", colour = NA),
      panel.background = element_rect(fill = "white", colour = NA),
      legend.position  = "bottom",
      plot.caption.position = "plot",
      axis.text.x = element_text(angle = 0, hjust = 0.5, vjust = 0.5)
    )
)

out_root <- file.path(out_dir, "Results_StringencyOnly26W")
out_xlsx_explor_26  <- file.path(out_root, "PerCountry_SI_LagSweep_0_26.xlsx")
out_xlsx_primary_26 <- file.path(out_root, "PerCountry_Primary_2to4.xlsx")
plot_dir_main <- file.path(out_root, "per_country_plots_0_26")
lags_explor <- 0:26
start_year <- 2020; start_week <- 1
end_year   <- 2022; end_week   <- 52
nw_lag     <- 8

dir.create(out_root,      showWarnings = FALSE, recursive = TRUE)
dir.create(plot_dir_main, showWarnings = FALSE, recursive = TRUE)

parse_year_week <- function(week_id_chr) {
  tibble(
    Week_Index = week_id_chr,
    Year = suppressWarnings(as.integer(stringr::str_sub(week_id_chr, 1, 4))),
    Week = suppressWarnings(as.integer(stringr::str_sub(week_id_chr, 6, 7)))
  )
}
nw_row <- function(mod, term, lag_nw = 8) {
  V  <- sandwich::NeweyWest(mod, lag = lag_nw, prewhite = FALSE, adjust = TRUE)
  ct <- lmtest::coeftest(mod, vcov. = V)
  if (!term %in% rownames(ct)) {
    return(tibble(term = term, estimate = NA_real_, std.error = NA_real_,
                  statistic = NA_real_, p.value = NA_real_))
  }
  tibble(
    term = term,
    estimate = unname(ct[term,1]),
    std.error = unname(ct[term,2]),
    statistic = unname(ct[term,3]),
    p.value = unname(ct[term,4])
  )
}

required_cols <- c(
  "CountryName","CountryCode","Year","Week","Week_Index",
  "Total_Deaths","Expected_Deaths","Excess_Deaths","Excess_percent",
  "Population","Deaths_per_100k","Excess_per_100k",
  "StringencyIndex_Average","GovernmentResponseIndex_Average",
  "ContainmentHealthIndex_Average","EconomicSupportIndex",
  "ConfirmedCases_weekly","ConfirmedDeaths_weekly",
  "COVID_Cases_per_100k","COVID_Deaths_per_100k"
)

df_raw <- read_xlsx(data_path)
miss <- setdiff(required_cols, names(df_raw))
if (length(miss) > 0) stop("Λείπουν στήλες από το Excel: ", paste(miss, collapse = ", "))

df2 <- df_raw %>%
  mutate(
    Week_Index = as.character(Week_Index),
    Year = suppressWarnings(as.integer(Year)),
    Week = suppressWarnings(as.integer(Week))
  )

if (any(is.na(df2$Year)) || any(is.na(df2$Week))) {
  df2 <- df2 %>% select(-Year, -Week) %>%
    left_join(parse_year_week(df2$Week_Index), by = "Week_Index")
}

df2 <- df2 %>%
  filter(
    (Year >  start_year | (Year == start_year & Week >= start_week)) &
      (Year <  end_year   | (Year == end_year   & Week <= end_week))
  ) %>%
  group_by(CountryCode, Week_Index) %>% slice(1L) %>% ungroup()

countries <- sort(unique(df2$CountryCode))

overview_list <- list()
sheets_explor <- list()

for (cc in countries) {
  d_cc <- df2 %>% filter(CountryCode == cc) %>%
    arrange(Year, Week) %>%
    mutate(SI = StringencyIndex_Average)
  
  lag_rows <- list()
  col_cc <- get_country_color(cc)
  
  for (k in lags_explor) {
    d_k <- d_cc %>%
      mutate(SI10_lag = dplyr::lag(SI, k) / 10) %>%
      filter(!is.na(Excess_per_100k), !is.na(SI10_lag))
    
    if (nrow(d_k) < 20) {
      lag_rows[[as.character(k)]] <- tibble(
        CountryCode = cc, lag = k,
        beta = NA_real_, se = NA_real_, t = NA_real_, p = NA_real_,
        conf_low = NA_real_, conf_high = NA_real_,
        N = nrow(d_k), R2 = NA_real_, AdjR2 = NA_real_
      )
      next
    }
    
    mod <- lm(Excess_per_100k ~ SI10_lag + factor(Week) + factor(Year), data = d_k)
    rob <- nw_row(mod, "SI10_lag", lag_nw = nw_lag)
    z <- qnorm(0.975)
    conf_low  <- rob$estimate - z * rob$std.error
    conf_high <- rob$estimate + z * rob$std.error
    gl <- broom::glance(mod)
    
    lag_rows[[as.character(k)]] <- tibble(
      CountryCode = cc, lag = k,
      beta = rob$estimate, se = rob$std.error, t = rob$statistic, p = rob$p.value,
      conf_low = conf_low, conf_high = conf_high,
      N = nobs(mod), R2 = gl$r.squared, AdjR2 = gl$adj.r.squared
    )
  }
  
  lag_tbl <- bind_rows(lag_rows) %>% arrange(lag) %>%
    mutate(p_fdr = p.adjust(p, method = "BH"))
  
  overview_list[[cc]] <- lag_tbl
  sheets_explor[[paste0("Lag_", cc)]] <- lag_tbl
  
  # --- Beta & CIs 
  p_beta <- ggplot(lag_tbl, aes(lag, beta)) +
    geom_hline(yintercept = 0, linetype = "dashed") +
    geom_point(color = col_cc) +
    geom_errorbar(aes(ymin = conf_low, ymax = conf_high), width = 0.15, color = col_cc) +
    scale_x_continuous(breaks = lags_explor) +
    labs(
      title = paste0(cc, ": Stringency/10 (lag) → Excess_per_100k"),
      subtitle = "FE(Week) + FE(Year), Newey–West SEs (lag=8) — Exploratory 0–26",
      x = "Lag (εβδομάδες)",
      y = "Beta (μονάδες Excess/100k ανά +10 μονάδες Stringency)"
    )
  ggsave(file.path(plot_dir_main, paste0("Lag_Beta_CI_", cc, "_0_26.png")),
         plot = p_beta, width = 8, height = 4.8, dpi = 300)
  
  # --- p-value ανά lag 
  p_pval <- ggplot(lag_tbl, aes(lag, p)) +
    geom_hline(yintercept = 0.05, linetype = "dashed") +
    geom_hline(yintercept = 0.01, linetype = "dotted") +
    geom_point(color = col_cc) +
    geom_line(alpha = 0.4, linewidth = 0.6, color = col_cc) +
    scale_x_continuous(breaks = lags_explor) +
    coord_cartesian(ylim = c(0, 1)) +
    labs(
      title = paste0(cc, ": p-value ανά lag (SI/10 → Excess/100k)"),
      subtitle = "Οριζόντιες: 0.05 (dashed), 0.01 (dotted)",
      x = "Lag (εβδομάδες)",
      y = "p-value"
    )
  ggsave(file.path(plot_dir_main, paste0("Lag_pvalue_", cc, "_0_26.png")),
         plot = p_pval, width = 8, height = 4.8, dpi = 300)
}

overview <- bind_rows(overview_list) %>%
  select(CountryCode, lag, beta, se, t, p, p_fdr, conf_low, conf_high, N, R2, AdjR2)

write_xlsx(c(list("Overview" = overview), sheets_explor), path = out_xlsx_explor_26)

# PRIMARY (2–4) 
primary_rows <- list()
for (cc in countries) {
  d_cc <- df2 %>% filter(CountryCode == cc) %>%
    arrange(Year, Week) %>%
    mutate(SI = StringencyIndex_Average)
  
  d_p <- d_cc %>%
    mutate(
      SI_lag2 = dplyr::lag(SI, 2),
      SI_lag3 = dplyr::lag(SI, 3),
      SI_lag4 = dplyr::lag(SI, 4),
      SI10_2to4 = rowMeans(cbind(SI_lag2, SI_lag3, SI_lag4), na.rm = FALSE) / 10
    ) %>%
    filter(!is.na(Excess_per_100k), !is.na(SI10_2to4))
  
  if (nrow(d_p) < 20) {
    primary_rows[[cc]] <- tibble(
      CountryCode = cc, beta = NA_real_, se = NA_real_, t = NA_real_, p = NA_real_,
      conf_low = NA_real_, conf_high = NA_real_, N = nrow(d_p),
      R2 = NA_real_, AdjR2 = NA_real_
    )
    next
  }
  
  mod_p <- lm(Excess_per_100k ~ SI10_2to4 + factor(Week) + factor(Year), data = d_p)
  rob_p <- nw_row(mod_p, "SI10_2to4", lag_nw = nw_lag)
  z <- qnorm(0.975)
  conf_low  <- rob_p$estimate - z * rob_p$std.error
  conf_high <- rob_p$estimate + z * rob_p$std.error
  gl <- broom::glance(mod_p)
  
  primary_rows[[cc]] <- tibble(
    CountryCode = cc,
    beta = rob_p$estimate, se = rob_p$std.error, t = rob_p$statistic, p = rob_p$p.value,
    conf_low = conf_low, conf_high = conf_high,
    N = nobs(mod_p), R2 = gl$r.squared, AdjR2 = gl$adj.r.squared
  )
}
primary_tbl <- bind_rows(primary_rows) %>% arrange(CountryCode)
write_xlsx(list("Primary_2to4" = primary_tbl), path = out_xlsx_primary_26)

message("✅ Έτοιμο (0–26, NW=8).",
        "\n- Exploratory: ", out_xlsx_explor_26,
        "\n- Primary (2–4): ", out_xlsx_primary_26,
        "\n- Plots: ", plot_dir_main, "\n")
