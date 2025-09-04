library(shiny)
library(readxl)
library(dplyr)
library(tidyr)
library(stringr)
library(gt)
library(scales)
library(ggtext)
library(ggplot2)
library(plotly)


library(DBI)
library(odbc)




file_path <- "Milliman's Financial Insight v4 (2).xlsx"
#wide_raw <- read_excel(file_path, sheet = "Part 2", col_names = FALSE, col_types = "text")
wide_raw <- readRDS("Snowflake.rds")

canon_period <- function(p) {
  p <- gsub("\\s+", "", toupper(as.character(p)))
  m <- stringr::str_match(p, "^(\\d{4})([QLT])(\\d+)$")
  if (any(is.na(m))) return(p)
  yr <- m[,2]; tag <- m[,3]; num <- m[,4]
  dplyr::case_when(
    tag %in% c("Q","L") ~ paste0(yr, "Q", num),
    tag == "T"         ~ paste0(yr, "H", num),
    TRUE               ~ p
  )
}

parse_period <- function(x) {
  x <- gsub("\\s+", "", toupper(as.character(x)))
  
  # Q/L/T → quarter/half
  m <- stringr::str_match(x, "^(\\d{4})([QLT])(\\d+)$")
  out <- rep(NA_character_, length(x))
  hit <- !is.na(m[,1])
  if (any(hit)) {
    yr <- m[hit,2]; tag <- m[hit,3]; num <- m[hit,4]
    out[hit] <- dplyr::case_when(
      tag %in% c("Q","L") ~ paste0(yr, "Q", num),
      tag == "T"          ~ paste0(yr, "H", num),
      TRUE                ~ x[hit]
    )
  }
  
  # YYYYMM → YYYYQ#
  mm <- stringr::str_match(x[!hit], "^(\\d{4})(0?[1-9]|1[0-2])$")
  if (!all(is.na(mm))) {
    yr2 <- mm[,2]; mon <- suppressWarnings(as.integer(mm[,3]))
    qtr <- paste0(yr2, "Q", ceiling(mon/3))
    out[!hit] <- ifelse(is.na(mon), x[!hit], qtr)
  }
  
  # fallback unchanged
  out[is.na(out)] <- x[is.na(out)]
  out
}


# Some Snowflake drivers keep the space in "Entity Name "
# Adjust this rename if your field is spelled a bit differently
df <- wide_raw %>%
  dplyr::rename(
    Entity    = `Entity Name `,
    PeriodRaw = SNLDATASOURCEPERIOD_3
  ) %>%
  dplyr::mutate(
    PeriodRaw = gsub("\\s+", "", toupper(PeriodRaw)),
    Quarter0  = parse_period(PeriodRaw),
    # normalize half-years to quarters if you want (optional):
    Quarter   = dplyr::if_else(grepl("^\\d{4}H[12]$", Quarter0),
                               sub("H1","Q1", sub("H2","Q3", Quarter0)),  # crude split
                               Quarter0
    ),
    Entity    = stringr::str_squish(stringr::str_trim(Entity))
  ) %>%
  dplyr::select(Entity, Quarter, dplyr::where(is.numeric)) %>%
  tidyr::pivot_longer(-c(Entity, Quarter), names_to = "Metric", values_to = "Value") %>%
  dplyr::filter(!is.na(Value)) %>%
  dplyr::mutate(Metric = stringr::str_to_upper(stringr::str_trim(Metric))) %>%
  # ensure uniqueness and collapse months→quarter
  dplyr::group_by(Entity, Quarter, Metric) %>%
  dplyr::summarise(Value = sum(Value, na.rm = TRUE), .groups = "drop")

company_list <- c(
  "Integris Insurance Co.", "Mutual RRG Inc.", "COPIC Insurance Co.",
  "Health Care Indemnity Inc.", "Medical Mutual Ins Co. of NC", "Physicians Insurance A Mult Co",
  "Ophthalmic Mult Ins Co (A RRG)", "Medical Protective Co", "ISMIE Mutual Insurance Co.",
  "Fortress Insurance Co.", "EmPRO Insurance Co.", "OMS National Insurance Co. RRG",
  "Medical Ins Exchange of CA", "Kansas Medical Mutual In Co", "Positive Physicians Ins Co.",
  "State Volunteer Mutual", "Mutual Insurance Co. of AZ", "The Doctors Co. an Interinsura",
  "CHART Indemnity Co.", "Medical Prof Mutual Ins Co.", "LAMMICO", "Medical Mutual Ins Co. of ME",
  "Medical Mutual Liab Ins Society", "Texas Builders Insurance Co.", "NCMIC Insurance Co.",
  "MAG Mutual Insurance Co.", "ProAssurance Indemnity Co.", "MCIC VT (A Reciprocal RRG)",
  "ProAssurance Ins Co. of Am", "NORCAL Ins Co.", "MLMIC Insurance Inc.",
  "UMIA Insurance Inc.", "MMIC Insurance Inc."
)

metrics_list <- df %>% distinct(Metric) %>% arrange(Metric) %>% pull(Metric)

all_entities <- sort(unique(df$Entity))

snl_pnc_group   <- all_entities[str_detect(all_entities, fixed("(SNL P&C Group)"))]
competitors     <- c(
  "Coverys (SNL P&C Group)", "Curi (SNL P&C Group)", "MagMutual (SNL P&C Group)",
  "MedPro Group Inc. (SNL P&C Subgroup)", "ProAssurance Corp. (SNL P&C Group)",
  "The Doctors Co. (SNL P&C Group)"
) %>% intersect(all_entities)


set.seed(42)
other_entities <- setdiff(all_entities, competitors)
random_cols     <- grDevices::rgb(runif(length(other_entities)), runif(length(other_entities)), runif(length(other_entities)))
entity_colors   <- c(
  setNames(random_cols, other_entities),
  "Coverys (SNL P&C Group)"         = "#8A3800",
  "Curi (SNL P&C Group)"            = "#00CED1",
  "MagMutual (SNL P&C Group)"       = "#1662FF",
  "MedPro Group Inc. (SNL P&C Subgroup)" = "#6928C5",
  "ProAssurance Corp. (SNL P&C Group)"   = "#198039",
  "The Doctors Co. (SNL P&C Group)"      = "#002C9C"
)

quarters_avail <- df$Quarter[grepl("^\\d{4}Q[1-4]$", df$Quarter)] |> unique() |> sort()
quarters_only <- df$Quarter[grepl("^\\d{4}Q[1-4]$", df$Quarter)] |> unique() |> sort()
default_a <- if (length(quarters_only) >= 2) quarters_only[length(quarters_only)-1] else NA
default_b <- if (length(quarters_only) >= 1) quarters_only[length(quarters_only)]     else NA


ui <- fluidPage(
  tags$head(
    tags$style(HTML(
      paste(
        c(
          ".horizontal-bar-graph { display: table; width: 100%; }",
          ".horizontal-bar-graph-segment { display: table-row; }",
          ".horizontal-bar-graph-label { display: table-cell; text-align: right; padding: 4px 10px 4px 0; white-space: nowrap; }",
          ".horizontal-bar-graph-value { display: table-cell; width: 100%; }",
          ".horizontal-bar-graph-value-bar { display: inline-block; height: 2em; position: relative; }",
          ".horizontal-bar-graph-value-bar:hover::after { content: attr(data-value); position: absolute; top: -1.5em; left: 50%; transform: translateX(-50%); background: #000; color: #fff; padding: 2px 6px; border-radius: 3px; }",
          ".no-data { color: red; font-style: italic; }"
        ),
        collapse = "\n"
      )
    ))
  ),
  titlePanel("Milliman Financial Dashboard"),
  tabsetPanel(
    tabPanel("Metric Compare",
             sidebarLayout(
               sidebarPanel(
                 selectInput("metric_comp", "Choose a Metric:", choices = metrics_list),
                 selectInput("period_a", "Period A:",
                             choices  = sort(unique(df$Quarter)),
                             selected = default_a
                 ),
                 selectInput("period_b", "Period B:",
                             choices  = sort(unique(df$Quarter)),
                             selected = default_b
                 )
               ),
               mainPanel(gt_output("table_cmp_all"), br())
             )
    ),
    tabPanel("Entities by Metric & Quarter",
             sidebarLayout(
               sidebarPanel(
                 checkboxInput("snl_pnc_only", "Show only SNL P&C Group", FALSE),
                 checkboxInput("hide_snl_pnc",   "Hide SNL P&C Group", FALSE),
                 radioButtons("mode", "Base Universe:", c("Competitors","All Entities"), "Competitors"),
                 selectInput("quarter", "Choose Quarter:", quarters_avail),
                 selectInput("metric",  "Choose Metric:",  choices = metrics_list),
                 uiOutput("entity_select_ui")
               ),
               mainPanel(
                 uiOutput("htmlBars")
               )
             )
    ),
    tabPanel("24 Q1 Vs 25 Q1 % Change",
             sidebarLayout(
               sidebarPanel(
                 selectInput("cmp_metric", "Choose a Metric:", choices = metrics_list),
                 selectInput("cmp_period_a", "Period A:",
                             choices  = quarters_only,
                             selected = default_a),
                 selectInput("cmp_period_b", "Period B:",
                             choices  = quarters_only,
                             selected = default_b)
               ),
               mainPanel(
                 plotlyOutput("comparison_plot", height = "400px"),
                 br(),
                 gt_output("comparison_tbl")
               )
             )
    )
    
  )
)

server <- function(input, output, session) {
  
  observe({
    if (length(quarters_only) >= 2) {
      updateSelectInput(session, "period_a",
                        choices = sort(unique(df$Quarter)),
                        selected = quarters_only[length(quarters_only)-1]
      )
      updateSelectInput(session, "period_b",
                        choices = sort(unique(df$Quarter)),
                        selected = quarters_only[length(quarters_only)]
      )
    }
  })
  df_compare_table <- reactive({
    req(input$metric_comp, input$period_a, input$period_b)
    pA <- canon_period(input$period_a); pB <- canon_period(input$period_b)
    
    w <- df %>%
      dplyr::filter(Metric == input$metric_comp, Entity %in% company_list) %>%
      tidyr::pivot_wider(names_from = Quarter, values_from = Value)
    
    if (!pA %in% names(w)) w[[pA]] <- NA_real_
    if (!pB %in% names(w)) w[[pB]] <- NA_real_
    
    w %>%
      dplyr::mutate(
        vA = .data[[pA]],
        vB = .data[[pB]],
        `Absolute Change` = vB - vA,
        `% Change` = dplyr::if_else(is.na(vA) | vA == 0, NA_real_, 100 * (vB - vA) / vA)
      ) %>%
      dplyr::select(Entity, vA, vB, `Absolute Change`, `% Change`)
  })
  
  
  output$table_cmp_all <- render_gt({
    dd <- df_compare_table()
    
    validate(
      need(nrow(dd) > 0, "No rows match this metric and cohort."),
      need(any(!is.na(dd$vA) | !is.na(dd$vB)), "No data in the selected periods for this metric and cohort.")
    )
    
    gt::gt(dd) %>%
      gt::cols_label(vA = input$period_a, vB = input$period_b,
                     `Absolute Change` = "Absolute Change", `% Change` = "% Change") %>%
      gt::fmt_number(columns = c(vA, vB, `Absolute Change`), decimals = 0) %>%
      gt::fmt_number(columns = `% Change`, decimals = 1, pattern = "{x}%") %>%
      gt::tab_header(title = input$metric_comp)
  })
  
  output$comparison_tbl <- render_gt({
    dd <- comp_df()
    
    validate(
      need(nrow(dd) > 0, "No rows match this metric and competitors set."),
      need(any(!is.na(dd$vA) | !is.na(dd$vB)), "No data in the selected periods for this metric and competitors.")
    )
    
    dd %>%
      dplyr::select(Entity, vA, vB, pct_change) %>%
      dplyr::mutate(`% Change` = round(pct_change, 1)) %>%
      gt::gt() %>%
      gt::cols_label(
        vA = input$cmp_period_a,
        vB = input$cmp_period_b,
        pct_change = "% Change"
      ) %>%
      gt::fmt_number(columns = c(vA, vB), decimals = 0) %>%
      gt::fmt_number(columns = pct_change, decimals = 1) %>%
      gt::tab_header(title = input$cmp_metric)
  })
  
  plot_df <- reactive({
    req(input$quarter, input$metric)
    uni <- if (input$snl_pnc_only) {
      snl_pnc_group
    } else if (input$hide_snl_pnc) {
      setdiff(all_entities, snl_pnc_group)
    } else if (input$mode == "Competitors") {
      competitors
    } else {
      all_entities
    }
    sel <- if (length(input$entities) > 0) input$entities else uni
    df %>%
      filter(
        Metric == input$metric,
        Quarter == input$quarter,
        Entity %in% sel
      ) %>%
      mutate(Entity = factor(Entity, levels = unique(Entity)))
  })
  
  output$htmlBars <- renderUI({
    d <- plot_df()
    if (nrow(d) == 0) {
      return(tags$p(class = 'no-data', 'No data for this selection.'))
    }
    
    segments <- lapply(seq_len(nrow(d)), function(i) {
      ent <- d$Entity[i]
      val <- d$Value[i]
      pct <- scales::rescale(val, to = c(0,100), from = range(d$Value, na.rm = TRUE))
      
      bar_col <- entity_colors[as.character(ent)]
      
      tags$div(
        class = 'horizontal-bar-graph-segment',
        tags$div(class = 'horizontal-bar-graph-label', ent),
        tags$div(
          class = 'horizontal-bar-graph-value',
          tags$span(
            class = 'horizontal-bar-graph-value-bar',
            style = sprintf('width:%.1f%%; background-color:%s;', pct, bar_col),
            `data-value` = formatC(val, format = 'f', digits = 1, big.mark = ',')
          )
        )
      )
    })
    
    do.call(tags$div, c(list(class = 'horizontal-bar-graph'), segments))
  })
  
  comp_df <- reactive({
    req(input$cmp_metric, input$cmp_period_a, input$cmp_period_b)
    pA <- canon_period(input$cmp_period_a)
    pB <- canon_period(input$cmp_period_b)
    
    w <- df %>%
      dplyr::filter(Metric == input$cmp_metric, Entity %in% competitors) %>%
      tidyr::pivot_wider(names_from = Quarter, values_from = Value)
    
    if (!pA %in% names(w)) w[[pA]] <- NA_real_
    if (!pB %in% names(w)) w[[pB]] <- NA_real_
    
    w %>%
      dplyr::mutate(
        vA = .data[[pA]],
        vB = .data[[pB]],
        abs_change = vB - vA,
        pct_change = dplyr::if_else(is.na(vA) | vA == 0, NA_real_, 100 * (vB - vA) / vA)
      )
  })
  
  output$comparison_plot <- renderPlotly({
    dd <- comp_df()
    pal <- entity_colors[names(entity_colors) %in% unique(dd$Entity)]
    if (!length(pal)) pal <- NULL
    
    if (all(is.na(dd$pct_change))) {
      # Fallback: show Absolute Change when % change is undefined
      p <- ggplot(dd, aes(
        x = Entity, y = abs_change, fill = Entity,
        text = paste0(
          "Entity: ", Entity,
          "<br>", input$cmp_period_a, ": ", round(vA, 0),
          "<br>", input$cmp_period_b, ": ", round(vB, 0),
          "<br>Abs Change: ", round(abs_change, 0)
        )
      )) +
        geom_col(show.legend = FALSE) +
        { if (is.null(pal)) scale_fill_discrete() else scale_fill_manual(values = pal) } +
        scale_y_continuous(name = "Absolute Change", expand = c(0,0)) +
        labs(x = NULL, title = paste(input$cmp_metric, "—", input$cmp_period_a, "→", input$cmp_period_b)) +
        theme_minimal() +
        theme(axis.text.x = element_text(angle = 45, vjust = 1, hjust = 1))
      
      return(ggplotly(p, tooltip = "text"))
    }
    
    # Normal % change plot
    p <- ggplot(dd, aes(
      x = Entity, y = pct_change, fill = Entity,
      text = paste0(
        "Entity: ", Entity,
        "<br>", input$cmp_period_a, ": ", round(vA, 0),
        "<br>", input$cmp_period_b, ": ", round(vB, 0),
        "<br>% Change: ", round(pct_change, 1), "%"
      )
    )) +
      geom_col(show.legend = FALSE) +
      { if (is.null(pal)) scale_fill_discrete() else scale_fill_manual(values = pal) } +
      scale_y_continuous(name = "% Change", expand = c(0,0)) +
      labs(x = NULL, title = paste(input$cmp_metric, "—", input$cmp_period_a, "→", input$cmp_period_b)) +
      theme_minimal() +
      theme(axis.text.x = element_text(angle = 45, vjust = 1, hjust = 1))
    
    ggplotly(p, tooltip = "text")
  })
  
  
  
  output$comparison_tbl <- render_gt({
    dd <- comp_df()
    
    validate(
      need(nrow(dd) > 0, "No rows match this metric and competitors set."),
      need(any(!is.na(dd$vA) | !is.na(dd$vB)), "No data in the selected periods for this metric and competitors.")
    )
    
    dd %>%
      dplyr::select(Entity, vA, vB, pct_change) %>%
      dplyr::mutate(`% Change` = round(pct_change, 1)) %>%
      gt::gt() %>%
      gt::cols_label(
        vA = input$cmp_period_a,
        vB = input$cmp_period_b,
        pct_change = "% Change"
      ) %>%
      gt::fmt_number(columns = c(vA, vB), decimals = 0) %>%
      gt::fmt_number(columns = pct_change, decimals = 1) %>%
      gt::tab_header(title = input$cmp_metric)
  })
  
  

}

shinyApp(ui, server)

