packages <- c(
  "quantmod", "dplyr", "tidyr", "lubridate", "zoo",
  "tseries", "urca", "strucchange", "fracdiff",
  "rugarch", "FinTS", "ggplot2", "knitr", "readxl",
  "moments", "writexl"
)

for (pkg in packages) {
  if (!requireNamespace(pkg, quietly = TRUE)) install.packages(pkg)
}
invisible(lapply(packages, library, character.only = TRUE))

set.seed(42)

print_table <- function(title, df, digits = 4) {
  cat("\n", strrep("=", 70), "\n", sep = "")
  cat(" ", title, "\n", sep = "")
  cat(strrep("=", 70), "\n", sep = "")
  print(knitr::kable(df, format = "simple", digits = digits))
  cat(strrep("-", 70), "\n\n", sep = "")
}

cat("\n[1/9] Collecting data...\n")

tickers_fi_infra <- c(
  "JURO11.SA", "KDIF11.SA", "CDII11.SA", "IFRI11.SA", "BDIF11.SA",
  "CPTI11.SA", "IFRA11.SA", "BODB11.SA", "BINC11.SA", "XPID11.SA",
  "BIDB11.SA", "RIFF11.SA", "SNID11.SA", "RBIF11.SA", "PRIF11.SA",
  "OGIN11.SA", "INFB11.SA", "BODI11.SA", "BRZD11.SA", "EXIF11.SA",
  "VANG11.SA", "IRIF11.SA", "ISNN11.SA", "ISNT11.SA", "ISEN11.SA",
  "ISET11.SA", "ISTT11.SA", "JMBI11.SA", "NUIF11.SA", "DBIN11.SA",
  "SUIN11.SA", "VALO11.SA", "VINF11.SA", "YVYC11.SA", "CPDI11.SA",
  "DIVS11.SA"
)

date_start <- "2019-01-01"
date_end   <- "2025-12-31"

raw_data <- list()
for (tk in tickers_fi_infra) {
  tryCatch({
    raw_data[[tk]] <- getSymbols(
      tk, src = "yahoo",
      from = date_start, to = date_end,
      auto.assign = FALSE, warnings = FALSE
    )
    cat("  OK:", tk, "\n")
  }, error = function(e) cat("  ERROR:", tk, "\n"))
}

ibov_raw <- getSymbols(
  "^BVSP", src = "yahoo",
  from = date_start, to = date_end,
  auto.assign = FALSE
)
ibov_df <- data.frame(
  Date  = zoo::index(ibov_raw),
  IBOV  = as.numeric(Ad(ibov_raw))
) |>
  filter(!is.na(IBOV)) |>
  mutate(Retorno_Log_IBOV = log(IBOV / lag(IBOV))) |>
  filter(!is.na(Retorno_Log_IBOV))

ifix_df <- read.csv2("IFIX.csv", stringsAsFactors = FALSE) |>
  rename(Date = 1, IFIX = 2) |>
  mutate(
    Date = as.Date(Date, format = "%d/%m/%Y"),
    IFIX = as.numeric(gsub(",", ".", IFIX))
  ) |>
  filter(!is.na(IFIX),
         Date >= as.Date(date_start),
         Date <= as.Date(date_end)) |>
  arrange(Date) |>
  mutate(Retorno_Log_IFIX = log(IFIX / lag(IFIX))) |>
  filter(!is.na(Retorno_Log_IFIX))

ida_df <- readxl::read_excel("IDA_IPCA.xlsx") |>
  rename(Date = 2, IDA_IPCA = 3) |>
  select(Date, IDA_IPCA) |>
  mutate(Date = as.Date(Date), IDA_IPCA = as.numeric(IDA_IPCA)) |>
  filter(!is.na(IDA_IPCA),
         Date >= as.Date(date_start),
         Date <= as.Date(date_end)) |>
  arrange(Date) |>
  mutate(Retorno_Log_IDA = log(IDA_IPCA / lag(IDA_IPCA))) |>
  filter(!is.na(Retorno_Log_IDA))

ida_ex_df <- readxl::read_xlsx("IDA_IPCA_ex_Infra.xlsx") |>
  rename(Date = 2, NumIndice = 3) |>
  select(Date, NumIndice) |>
  mutate(Date = as.Date(Date), NumIndice = as.numeric(NumIndice)) |>
  filter(!is.na(NumIndice),
         Date >= as.Date(date_start),
         Date <= as.Date(date_end)) |>
  arrange(Date) |>
  mutate(Retorno_Log_IDA_ex = log(NumIndice / lag(NumIndice))) |>
  filter(!is.na(Retorno_Log_IDA_ex))

cat("  Data collection complete.\n")

cat("\n[2/9] Computing liquidity metrics...\n")

calc_adtv <- function(tk) {
  tryCatch({
    d   <- raw_data[[tk]]; if (is.null(d)) return(NA)
    vol <- grep("Volume",   colnames(d), value = TRUE)
    adj <- grep("Adjusted", colnames(d), value = TRUE)[1]
    if (length(vol) == 0 || length(adj) == 0) return(NA)
    mean(as.numeric(d[, vol]) * as.numeric(d[, adj]), na.rm = TRUE)
  }, error = function(e) NA)
}

calc_freq <- function(tk) {
  tryCatch({
    d   <- raw_data[[tk]]; if (is.null(d)) return(NA)
    adj <- grep("Adjusted", colnames(d), value = TRUE)[1]
    mean(!is.na(as.numeric(d[, adj])) & as.numeric(d[, adj]) > 0)
  }, error = function(e) NA)
}

calc_nobs <- function(tk) {
  tryCatch({
    d   <- raw_data[[tk]]; if (is.null(d)) return(0)
    adj <- grep("Adjusted", colnames(d), value = TRUE)[1]
    sum(!is.na(as.numeric(d[, adj])))
  }, error = function(e) 0)
}

liquidity_metrics <- data.frame(
  Ticker     = names(raw_data),
  ADTV       = sapply(names(raw_data), calc_adtv),
  Frequencia = sapply(names(raw_data), calc_freq),
  N_Obs      = sapply(names(raw_data), calc_nobs),
  stringsAsFactors = FALSE
)

selected_tickers <- liquidity_metrics |>
  filter(!is.na(ADTV), ADTV >= 50000, Frequencia >= 0.60, N_Obs >= 250) |>
  pull(Ticker)

table_2 <- liquidity_metrics |>
  filter(Ticker %in% selected_tickers) |>
  arrange(desc(ADTV))

print_table("TABLE 2 — Liquidity Metrics of Selected FI-Infra Funds", table_2)

cat("\n[3/9] Building the index...\n")

prices_list <- lapply(selected_tickers, function(tk) {
  d   <- raw_data[[tk]]
  col <- grep("Adjusted", colnames(d), value = TRUE)[1]
  obj <- d[, col]
  colnames(obj) <- tk
  obj
})
prices_xts  <- do.call(merge, prices_list)
prices_locf <- na.locf(prices_xts, na.rm = FALSE)
returns_xts <- diff(log(prices_locf))[-1, ]

adtv_sel <- liquidity_metrics |>
  filter(Ticker %in% selected_tickers) |>
  pull(ADTV)
names(adtv_sel) <- selected_tickers

weights <- adtv_sel / sum(adtv_sel, na.rm = TRUE)
for (i in 1:20) {
  over_cap <- weights > 0.15
  if (!any(over_cap)) break
  weights[over_cap]  <- 0.15
  weights[!over_cap] <- weights[!over_cap] *
    ((1 - sum(over_cap) * 0.15) / sum(weights[!over_cap]))
}

cat("  Final weights (after 15% cap):\n")
print(round(sort(weights, decreasing = TRUE), 4))

ret_mat <- coredata(returns_xts)
iia_ret_vec <- apply(ret_mat, 1, function(row) {
  valid <- !is.na(row)
  if (sum(valid) == 0) return(NA)
  sum(row[valid] * (weights[valid] / sum(weights[valid])))
})
iia_ret <- xts::xts(iia_ret_vec, order.by = zoo::index(returns_xts))
cat("  IIA-INFRA built:", length(iia_ret), "observations.\n")

cat("\n[4/9] Aligning calendar to IBOV (B3 trading days)...\n")

master_dates <- data.frame(Date = ibov_df$Date)

series_xts <- list(
  IIA_INFRA = iia_ret,
  IFIX      = xts::xts(ifix_df$Retorno_Log_IFIX,     order.by = ifix_df$Date),
  IBOV      = xts::xts(ibov_df$Retorno_Log_IBOV,     order.by = ibov_df$Date),
  IDA_IPCA  = xts::xts(ida_df$Retorno_Log_IDA,       order.by = ida_df$Date),
  IDA_ex    = xts::xts(ida_ex_df$Retorno_Log_IDA_ex, order.by = ida_ex_df$Date)
)

df_master <- master_dates
for (nm in names(series_xts)) {
  tmp <- data.frame(
    Date = zoo::index(series_xts[[nm]]),
    Val  = as.numeric(series_xts[[nm]])
  )
  df_master        <- left_join(df_master, tmp, by = "Date")
  df_master$Val[is.na(df_master$Val)] <- 0
  colnames(df_master)[ncol(df_master)] <- nm
}

series_names <- names(series_xts)

cat(
  "  Period:", format(min(df_master$Date), "%d/%m/%Y"),
  "to",       format(max(df_master$Date), "%d/%m/%Y"),
  "| N =", nrow(df_master), "observations\n"
)

cat("\n[5/9] Computing descriptive statistics and unit-root tests...\n")

table_4 <- do.call(rbind, lapply(series_names, function(nm) {
  x  <- df_master[[nm]] * 100
  jb <- moments::jarque.test(x)
  data.frame(
    Serie    = nm,
    N        = length(x),
    Media    = mean(x),
    Mediana  = median(x),
    DP       = sd(x),
    Skewness = moments::skewness(x),
    ExKurt   = moments::kurtosis(x) - 3,
    Min      = min(x),
    Max      = max(x),
    JB_pval  = jb$p.value,
    stringsAsFactors = FALSE
  )
}))

print_table(
  "TABLE 4 — Descriptive Statistics of Daily Log-Returns (%)",
  table_4
)

ur_results <- list()
for (nm in series_names) {
  x <- df_master[[nm]]
  
  for (spec in c("drift", "trend", "none")) {
    adf <- urca::ur.df(x, type = spec, selectlags = "BIC")
    ur_results[[paste(nm, "ADF", spec)]] <- data.frame(
      Serie      = nm,
      Teste      = "ADF",
      Spec       = spec,
      Estat      = round(adf@teststat[1], 3),
      VC_5pct    = adf@cval[1, "5pct"],
      Rejeita_H0 = ifelse(adf@teststat[1] < adf@cval[1, "5pct"], "Sim", "Não")
    )
  }
  
  pp <- urca::ur.pp(x, type = "Z-tau", model = "constant", lags = "short")
  ur_results[[paste(nm, "PP")]] <- data.frame(
    Serie      = nm,
    Teste      = "PP",
    Spec       = "c",
    Estat      = round(pp@teststat[1], 3),
    VC_5pct    = pp@cval[1, "5pct"],
    Rejeita_H0 = ifelse(pp@teststat[1] < pp@cval[1, "5pct"], "Sim", "Não")
  )
}

table_5 <- do.call(rbind, ur_results)
print_table("TABLE 5 — Unit-Root Tests: ADF and Phillips-Perron", table_5)

kpss_results <- list()
for (nm in series_names) {
  x <- df_master[[nm]]
  for (spec in c("mu", "tau")) {
    kpss <- urca::ur.kpss(x, type = spec, lags = "short")
    kpss_results[[paste(nm, spec)]] <- data.frame(
      Serie      = nm,
      Spec       = spec,
      Estat      = round(kpss@teststat[1], 4),
      VC_5pct    = kpss@cval[1, "5pct"],
      Rejeita_H0 = ifelse(kpss@teststat[1] > kpss@cval[1, "5pct"], "Sim", "Não")
    )
  }
}

table_6 <- do.call(rbind, kpss_results)
print_table("TABLE 6 — KPSS Stationarity Test", table_6)

cat("\n[6/9] Identifying structural breaks (Bai-Perron)...\n")

break_dates_map <- list()
bp_results      <- list()

for (nm in series_names) {
  x  <- df_master[[nm]]
  dx <- df_master$Date
  
  tryCatch({
    bp     <- strucchange::breakpoints(x^2 ~ 1, breaks = 5, h = 0.15)
    bp_opt <- strucchange::breakpoints(bp)
    
    if (!is.na(bp_opt$breakpoints[1])) {
      break_dates_map[[nm]] <- dx[bp_opt$breakpoints]
      bp_results[[nm]] <- data.frame(
        Serie     = nm,
        N_Quebras = length(bp_opt$breakpoints),
        Datas     = paste(format(break_dates_map[[nm]], "%d/%m/%Y"), collapse = "; ")
      )
    } else {
      break_dates_map[[nm]] <- as.Date(character(0))
      bp_results[[nm]] <- data.frame(Serie = nm, N_Quebras = 0, Datas = "None")
    }
  }, error = function(e) {
    break_dates_map[[nm]] <<- as.Date(character(0))
    bp_results[[nm]]      <<- data.frame(Serie = nm, N_Quebras = 0, Datas = "Error")
  })
}

table_7 <- do.call(rbind, bp_results)
print_table("TABLE 7 — Structural Breaks (Bai and Perron, 2003)", table_7)

cat("\n[7/9] Estimating long-memory parameter d (GSEM — Robinson 1995)...\n")

gsem <- function(x, m) {
  spec <- spec.pgram(x, taper = 0, plot = FALSE)
  freq <- spec$freq[1:m]
  Iw   <- spec$spec[1:m]
  
  objective <- function(d) {
    G <- mean(((2 * pi * freq)^(2 * d)) * Iw)
    if (G <= 0) return(Inf)
    log(G) - 2 * d * mean(log(2 * pi * freq))
  }
  
  opt <- optimize(objective, interval = c(-0.49, 0.49))
  se  <- 1 / (2 * sqrt(m))
  
  list(
    d     = opt$minimum,
    se    = se,
    ic_lo = opt$minimum - 1.96 * se,
    ic_hi = opt$minimum + 1.96 * se
  )
}

gsem_results <- list()
for (nm in series_names) {
  x <- as.numeric(df_master[[nm]])
  n <- length(x)
  
  r05 <- tryCatch(gsem(x, floor(n^0.5)), error = function(e) NULL)
  r06 <- tryCatch(gsem(x, floor(n^0.6)), error = function(e) NULL)
  r07 <- tryCatch(gsem(x, floor(n^0.7)), error = function(e) NULL)
  
  if (!is.null(r06)) {
    conclusion <- ifelse(
      r06$ic_lo > 0, "Long memory (d > 0)",
      ifelse(r06$ic_hi < 0, "Anti-persistence", "H0 not rejected")
    )
    gsem_results[[nm]] <- data.frame(
      Serie     = nm,
      d_m05     = ifelse(!is.null(r05), round(r05$d, 4), NA),
      d_m06     = round(r06$d, 4),
      d_m07     = ifelse(!is.null(r07), round(r07$d, 4), NA),
      IC_Inf    = round(r06$ic_lo, 4),
      IC_Sup    = round(r06$ic_hi, 4),
      Conclusao = conclusion,
      stringsAsFactors = FALSE
    )
  }
}

table_8 <- do.call(rbind, gsem_results)
print_table(
  "TABLE 8 — Semiparametric Long-Memory Estimation: d (GSEM)",
  table_8
)

cat("\n[8/9] Estimating ARMA-FIGARCH models (this may take a few minutes)...\n")

fitted_models  <- list()
coef_results   <- list()
spec_results   <- list()
diag_results   <- list()

for (nm in series_names) {
  cat("  Estimating:", nm, "... ")
  x  <- df_master[[nm]]
  dx <- df_master$Date
  dq <- break_dates_map[[nm]]
  
  dummies <- NULL
  if (length(dq) > 0) {
    dummies <- matrix(0, nrow = length(x), ncol = length(dq))
    for (i in seq_along(dq)) dummies[dx == dq[i], i] <- 1
  }
  
  spec_figarch <- rugarch::ugarchspec(
    variance.model     = list(model = "fiGARCH", garchOrder = c(1, 1)),
    mean.model         = list(armaOrder = c(1, 1), include.mean = TRUE,
                              external.regressors = dummies),
    distribution.model = "std"
  )
  spec_garch <- rugarch::ugarchspec(
    variance.model     = list(model = "sGARCH", garchOrder = c(1, 1)),
    mean.model         = list(armaOrder = c(1, 1), include.mean = TRUE,
                              external.regressors = dummies),
    distribution.model = "std"
  )
  
  fit_result <- tryCatch({
    f <- rugarch::ugarchfit(spec_figarch, x, solver = "hybrid")
    if (rugarch::convergence(f) == 0) {
      list(fit = f, model = "FIGARCH")
    } else {
      stop("Did not converge")
    }
  }, error = function(e) {
    tryCatch({
      f2 <- rugarch::ugarchfit(spec_garch, x, solver = "hybrid")
      list(fit = f2, model = "GARCH [Fallback]")
    }, error = function(e2) NULL)
  })
  
  if (!is.null(fit_result)) {
    cat(fit_result$model, "\n")
    fitted_models[[nm]] <- fit_result$fit
    spec_results[[nm]]  <- data.frame(
      Serie  = nm,
      Equacao = "ARMA(1,1)",
      Modelo  = fit_result$model
    )
    
    cf <- as.data.frame(fit_result$fit@fit$matcoef)
    cf$Parametro <- rownames(cf)
    cf$Serie     <- nm
    rownames(cf) <- NULL
    coef_results[[nm]] <- cf[, c("Serie", "Parametro",
                                 " Estimate", " Std. Error", "Pr(>|t|)")]
    
    z_std  <- as.numeric(rugarch::residuals(fit_result$fit, standardize = TRUE))
    lb_z   <- Box.test(z_std,    lag = 20, type = "Ljung-Box")
    lb_z2  <- Box.test(z_std^2,  lag = 20, type = "Ljung-Box")
    arch_t <- FinTS::ArchTest(z_std, lags = 10)
    
    diag_results[[nm]] <- data.frame(
      Serie     = nm,
      LB_pval   = round(lb_z$p.value,   4),
      LB2_pval  = round(lb_z2$p.value,  4),
      ARCH_pval = round(arch_t$p.value, 4),
      Conclusao = ifelse(
        lb_z$p.value > 0.05 & arch_t$p.value > 0.05,
        "Adequate", "Inadequate"
      )
    )
  } else {
    cat("FAILED\n")
  }
}

table_9       <- do.call(rbind, spec_results)
table_9b_coef <- do.call(rbind, coef_results)
table_10      <- do.call(rbind, diag_results)

print_table("TABLE 9 — Model Specification per Series", table_9)

cat("\n", strrep("=", 70), "\n", sep = "")
cat(" TABLE 9B — Estimated Parameters: ARMA-FIGARCH Models\n")
cat(strrep("=", 70), "\n", sep = "")
for (nm in names(coef_results)) {
  cat("\n  >>", nm, "\n")
  print(knitr::kable(coef_results[[nm]], format = "simple", digits = 6))
}
cat(strrep("-", 70), "\n\n")

print_table(
  "TABLE 10 — Residual Diagnostics: Ljung-Box and ARCH-LM",
  table_10
)

cat("\n[9/9] Computing correlations and running backtests...\n")

std_resid_mat <- do.call(
  cbind,
  lapply(fitted_models, function(f)
    as.numeric(rugarch::residuals(f, standardize = TRUE)))
)

cor_matrix <- round(cor(std_resid_mat), 4)
cat("\n", strrep("=", 70), "\n", sep = "")
cat(" TABLE 11A — Full Pearson Correlation Matrix (Standardised Residuals)\n")
cat(strrep("=", 70), "\n", sep = "")
print(knitr::kable(as.data.frame(cor_matrix), format = "simple", digits = 4))
cat(strrep("-", 70), "\n\n")

benchmark_names <- setdiff(colnames(std_resid_mat), "IIA_INFRA")
table_11 <- data.frame(
  Par        = paste("IIA_INFRA x", benchmark_names),
  Correlacao = sapply(benchmark_names, function(nm2)
    round(cor(std_resid_mat[, "IIA_INFRA"], std_resid_mat[, nm2]), 4)),
  stringsAsFactors = FALSE
)
print_table(
  "TABLE 11B — Pearson Correlations: IIA-INFRA vs Benchmarks",
  table_11
)

christoffersen_ind <- function(It) {
  n <- length(It)
  if (n < 2) return(list(lr_ind = NA, pval_ind = NA))
  
  n00 <- sum(It[-n] == 0 & It[-1] == 0)
  n01 <- sum(It[-n] == 0 & It[-1] == 1)
  n10 <- sum(It[-n] == 1 & It[-1] == 0)
  n11 <- sum(It[-n] == 1 & It[-1] == 1)
  
  pi01 <- ifelse((n00 + n01) > 0, n01 / (n00 + n01), 0)
  pi11 <- ifelse((n10 + n11) > 0, n11 / (n10 + n11), 0)
  pi_h <- (n01 + n11) / (n00 + n01 + n10 + n11)
  
  if (any(c(pi01, pi11, pi_h) <= 0) || any(c(pi01, pi11, pi_h) >= 1))
    return(list(lr_ind = NA, pval_ind = NA))
  
  lr_ind <- -2 * (
    (n00 + n10) * log(1 - pi_h) + (n01 + n11) * log(pi_h) -
      n00 * log(1 - pi01) - n01 * log(pi01) -
      n10 * log(1 - pi11) - n11 * log(pi11)
  )
  list(lr_ind = lr_ind, pval_ind = pchisq(lr_ind, df = 1, lower.tail = FALSE))
}

backtest_start <- as.Date("2024-01-01")
bt_results     <- list()

for (nm in names(fitted_models)) {
  f   <- fitted_models[[nm]]
  dx  <- df_master$Date
  idx <- which(dx >= backtest_start)
  if (length(idx) < 50) next
  
  ret_empirical <- as.numeric(df_master[[nm]])[idx]
  mu_hat        <- as.numeric(rugarch::fitted(f))[idx]
  sigma_hat     <- as.numeric(rugarch::sigma(f))[idx]
  nu            <- coef(f)["shape"]
  
  for (alpha in c(0.95, 0.99)) {
    p_tail <- 1 - alpha
    
    var_t <- -(mu_hat + sigma_hat *
                 rugarch::qdist("std", p = p_tail,
                                mu = 0, sigma = 1, shape = nu))
    
    It  <- as.integer(ret_empirical < -abs(var_t))
    N   <- sum(It)
    T_n <- length(idx)
    ph  <- N / T_n
    
    pval_kupiec <- NA
    if (ph > 0 && ph < 1) {
      lr_uc <- -2 * (
        N * log(p_tail) + (T_n - N) * log(1 - p_tail) -
          N * log(ph)    - (T_n - N) * log(1 - ph)
      )
      pval_kupiec <- pchisq(lr_uc, df = 1, lower.tail = FALSE)
    }
    
    chris <- christoffersen_ind(It)
    
    conclusion <- dplyr::case_when(
      is.na(pval_kupiec)                                    ~ "Inconclusive",
      pval_kupiec < 0.05                                    ~ "Rejects Coverage",
      !is.na(chris$pval_ind) & chris$pval_ind < 0.05       ~ "Rejects Independence",
      TRUE                                                  ~ "Adequate"
    )
    
    bt_results[[paste(nm, alpha)]] <- data.frame(
      Serie          = nm,
      Nivel          = paste0(alpha * 100, "%"),
      Viol_Obs       = N,
      Viol_Esp       = round(T_n * p_tail, 1),
      Razao_Viol     = round(ph / p_tail, 3),
      Kupiec_pval    = ifelse(!is.na(pval_kupiec), round(pval_kupiec, 4), NA),
      Chris_IND_pval = ifelse(!is.na(chris$pval_ind), round(chris$pval_ind, 4), NA),
      Conclusao      = conclusion,
      stringsAsFactors = FALSE
    )
  }
}

table_12 <- do.call(rbind, bt_results)
print_table(
  "TABLE 12 — VaR Backtesting: Kupiec (1995) and Christoffersen (1998)",
  table_12
)

cat("\nExporting results...\n")

export_list <- list(
  "Tab2_Liquidez"     = table_2,
  "Tab4_Descritivas"  = table_4,
  "Tab5_ADF_PP"       = table_5,
  "Tab6_KPSS"         = table_6,
  "Tab7_Bai_Perron"   = table_7,
  "Tab8_GSEM"         = table_8,
  "Tab9_Modelos"      = table_9,
  "Tab9b_Coefs_GARCH" = table_9b_coef,
  "Tab10_Diagnostico" = table_10,
  "Tab11_Correlacao"  = table_11,
  "Tab12_Backtesting" = table_12
)

writexl::write_xlsx(export_list, path = "ResultadosR.xlsx")

cat("\n", strrep("=", 70), "\n", sep = "")
cat(" SUCCESS — All tables printed to console and exported.\n")
cat(" Output file: Resultados.xlsx\n")
cat(strrep("=", 70), "\n\n", sep = "")