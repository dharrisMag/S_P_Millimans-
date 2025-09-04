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


# file_path <- "Milliman's Financial Insight v4 (2).xlsx"
# wide_raw <- read_excel(file_path, sheet = "Part 2", col_names = TRUE)
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
  mm <- stringr::str_match(x[!hit], "^(\\d{4})(0?[1-9]|1[0-2])$")
  if (!all(is.na(mm))) {
    yr2 <- mm[,2]; mon <- suppressWarnings(as.integer(mm[,3]))
    qtr <- paste0(yr2, "Q", ceiling(mon/3))
    out[!hit] <- ifelse(is.na(mon), x[!hit], qtr)
  }
  out[is.na(out)] <- x[is.na(out)]
  out
}



entity_col <- if ("Entity Name " %in% names(wide_raw)) "Entity Name " else {
  nm <- names(wide_raw)[str_detect(names(wide_raw), regex("^Entity$", ignore_case = TRUE))]
  if (length(nm) == 0) stop("Couldn't find an Entity column.")
  nm[1]
}
period_col <- if ("SNLDATASOURCEPERIOD_3" %in% names(wide_raw)) "SNLDATASOURCEPERIOD_3" else {
  nm <- names(wide_raw)[str_detect(names(wide_raw), regex("SNLDATASOURCEPERIOD_3", ignore_case = TRUE))]
  if (length(nm) == 0) stop("Couldn't find SNLDATASOURCEPERIOD_3.")
  nm[1]
}

df <- wide_raw %>%
  mutate(
    Entity    = stringr::str_squish(stringr::str_trim(.data[[entity_col]])),
    PeriodRaw = .data[[period_col]]
  ) %>%
  mutate(
    PeriodRaw = gsub("\\s+", "", toupper(PeriodRaw)),
    Quarter0  = parse_period(PeriodRaw),
    Quarter   = dplyr::if_else(
      grepl("^\\d{4}H[12]$", Quarter0),
      sub("H1","Q1", sub("H2","Q3", Quarter0)),   # split halves crudely into starting quarters
      Quarter0
    )
  ) %>%
  select(Entity, Quarter, where(is.numeric)) %>%
  pivot_longer(-c(Entity, Quarter), names_to = "Metric", values_to = "Value") %>%
  filter(!is.na(Value)) %>%
  mutate(Metric = stringr::str_to_upper(stringr::str_trim(Metric))) %>%
  group_by(Entity, Quarter, Metric) %>%
  summarise(Value = sum(Value, na.rm = TRUE), .groups = "drop")


all_entities  <- sort(unique(df$Entity))
metrics_list  <- df %>% distinct(Metric) %>% arrange(Metric) %>% pull(Metric)
quarters_only <- df$Quarter[grepl("^\\d{4}Q[1-4]$", df$Quarter)] |> unique() |> sort()
quarters_avail <- quarters_only
default_a <- if (length(quarters_only) >= 2) quarters_only[length(quarters_only)-1] else quarters_only[1]
default_b <- if (length(quarters_only) >= 1) quarters_only[length(quarters_only)] else quarters_only[1]


competitor_patterns <- c(
  "Coverys",
  "Curi",
  "Mag\\s*Mutual|MagMutual|MAG Mutual",
  "MedPro|Medical\\s*Protective",
  "ProAssurance",
  "The\\s*Doctors\\s*Co|TDC"
)
competitors_present <- all_entities[sapply(all_entities, function(e)
  any(stringr::str_detect(e, regex(paste(competitor_patterns, collapse="|"), ignore_case = TRUE)))
)]
# Fallback to SNL tags if regex found none
if (length(competitors_present) == 0) {
  competitors_present <- all_entities[grepl("\\(SNL P&C (Group|Subgroup)\\)", all_entities)]
}

# Colors
set.seed(42)
random_cols   <- grDevices::rgb(runif(length(all_entities)), runif(length(all_entities)), runif(length(all_entities)))
entity_colors <- setNames(random_cols, all_entities)
entity_colors[names(entity_colors) %>% stringr::str_detect(regex("Coverys", ignore_case=TRUE))] <- "#8A3800"
entity_colors[names(entity_colors) %>% stringr::str_detect(regex("Curi", ignore_case=TRUE))]    <- "#00CED1"
entity_colors[names(entity_colors) %>% stringr::str_detect(regex("Mag\\s*Mutual|MagMutual", ignore_case=TRUE))] <- "#1662FF"
entity_colors[names(entity_colors) %>% stringr::str_detect(regex("MedPro|Medical\\s*Protective", ignore_case=TRUE))] <- "#6928C5"
entity_colors[names(entity_colors) %>% stringr::str_detect(regex("ProAssurance", ignore_case=TRUE))] <- "#198039"
entity_colors[names(entity_colors) %>% stringr::str_detect(regex("The\\s*Doctors\\s*Co|TDC", ignore_case=TRUE))] <- "#002C9C"

# ============== UI ==============
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
                 selectInput("period_a", "Period A:", choices = quarters_avail, selected = default_a),
                 selectInput("period_b", "Period B:", choices = quarters_avail, selected = default_b)
               ),
               mainPanel(gt_output("table_cmp_all"), br())
             )
    ),
    tabPanel("Entities by Metric & Quarter",
             sidebarLayout(
               sidebarPanel(
                 radioButtons("mode", "Base Universe:", c("Competitors","All Entities"), "Competitors"),
                 checkboxInput("snl_pnc_only", "Show only SNL Group labels", FALSE),
                 checkboxInput("hide_snl_pnc", "Hide SNL Group labels", FALSE),
                 selectInput("quarter", "Choose Quarter:", quarters_avail, selected = default_b),
                 selectInput("metric",  "Choose Metric:",  choices = metrics_list),
                 uiOutput("entity_select_ui")
               ),
               mainPanel(uiOutput("htmlBars"))
             )
    ),
    tabPanel("24 Q1 Vs 25 Q1 % Change",
             sidebarLayout(
               sidebarPanel(
                 selectInput("cmp_metric", "Choose a Metric:", choices = metrics_list),
                 selectInput("cmp_period_a", "Period A:", choices = quarters_avail, selected = default_a),
                 selectInput("cmp_period_b", "Period B:", choices = quarters_avail, selected = default_b)
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
  
  # Keep latest two quarters selected by default as data evolves
  observe({
    if (length(quarters_only) >= 2) {
      updateSelectInput(session, "period_a", choices = quarters_avail, selected = quarters_only[length(quarters_only)-1])
      updateSelectInput(session, "period_b", choices = quarters_avail, selected = quarters_only[length(quarters_only)])
      updateSelectInput(session, "cmp_period_a", choices = quarters_avail, selected = quarters_only[length(quarters_only)-1])
      updateSelectInput(session, "cmp_period_b", choices = quarters_avail, selected = quarters_only[length(quarters_only)])
    }
  })
  

  df_compare_table <- reactive({
    req(input$metric_comp, input$period_a, input$period_b)
    pA <- canon_period(input$period_a); pB <- canon_period(input$period_b)
    
    w <- df %>%
      filter(Metric == input$metric_comp) %>%                 # no hard-coded company list
      pivot_wider(names_from = Quarter, values_from = Value)
    
    if (!pA %in% names(w)) w[[pA]] <- NA_real_
    if (!pB %in% names(w)) w[[pB]] <- NA_real_
    
    w %>%
      mutate(
        vA = .data[[pA]],
        vB = .data[[pB]],
        rank_val = dplyr::coalesce(vB, vA, 0),
        `Absolute Change` = vB - vA,
        `% Change` = if_else(is.na(vA) | vA == 0, NA_real_, 100 * (vB - vA) / vA)
      ) %>%
      arrange(desc(rank_val)) %>%
      slice_head(n = 15) %>%
      select(Entity, vA, vB, `Absolute Change`, `% Change`)
  })
  
  output$table_cmp_all <- render_gt({
    dd <- df_compare_table()
    validate(
      need(nrow(dd) > 0, "No rows available for this metric."),
      need(any(!is.na(dd$vA) | !is.na(dd$vB)), "No data in the selected periods for this metric.")
    )
    gt(dd) %>%
      cols_label(vA = input$period_a, vB = input$period_b) %>%
      fmt_number(columns = c(vA, vB, `Absolute Change`), decimals = 0) %>%
      fmt_number(columns = `% Change`, decimals = 1, pattern = "{x}%") %>%
      tab_header(title = paste0(input$metric_comp, " — Top 15 by ", input$period_b))
  })
  

  universe_entities <- reactive({
    base <- if (input$mode == "Competitors") competitors_present else all_entities
    if (input$snl_pnc_only) {
      base <- base[grepl("\\(SNL P&C (Group|Subgroup)\\)", base)]
    }
    if (input$hide_snl_pnc) {
      base <- base[!grepl("\\(SNL P&C (Group|Subgroup)\\)", base)]
    }
    sort(unique(base))
  })
  
  output$entity_select_ui <- renderUI({
    uni <- universe_entities()
    selectizeInput("entities", "Entities (optional):",
                   choices = uni, selected = uni,
                   multiple = TRUE,
                   options = list(plugins = list("remove_button"))
    )
  })
  
  plot_df <- reactive({
    req(input$quarter, input$metric)
    uni <- universe_entities()
    sel <- if (!is.null(input$entities) && length(input$entities) > 0) input$entities else uni
    
    df %>%
      filter(Metric == input$metric, Quarter == input$quarter, Entity %in% sel) %>%
      arrange(desc(Value)) %>%
      mutate(Entity = factor(Entity, levels = Entity))
  })
  
  output$htmlBars <- renderUI({
    d <- plot_df()
    if (nrow(d) == 0) return(tags$p(class = 'no-data', 'No data for this selection.'))
    
    rng <- range(d$Value, na.rm = TRUE)
    segments <- lapply(seq_len(nrow(d)), function(i) {
      ent <- d$Entity[i]
      val <- d$Value[i]
      pct <- if (diff(rng) > 0) scales::rescale(val, to = c(0,100), from = rng) else 100
      bar_col <- entity_colors[as.character(ent)]
      tags$div(
        class = 'horizontal-bar-graph-segment',
        tags$div(class = 'horizontal-bar-graph-label', as.character(ent)),
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
      filter(Metric == input$cmp_metric, Entity %in% competitors_present) %>%
      pivot_wider(names_from = Quarter, values_from = Value)
    
    if (!pA %in% names(w)) w[[pA]] <- NA_real_
    if (!pB %in% names(w)) w[[pB]] <- NA_real_
    
    w %>%
      mutate(
        vA = .data[[pA]],
        vB = .data[[pB]],
        abs_change = vB - vA,
        pct_change = if_else(is.na(vA) | vA == 0, NA_real_, 100 * (vB - vA) / vA)
      )
  })
  
  output$comparison_plot <- renderPlotly({
    dd <- comp_df()
    validate(need(nrow(dd) > 0, "No competitor rows available for this metric/periods."))
    pal <- entity_colors[names(entity_colors) %in% unique(dd$Entity)]
    if (!length(pal)) pal <- NULL
    
    if (all(is.na(dd$pct_change))) {
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
      select(Entity, vA, vB, pct_change) %>%
      mutate(`% Change` = round(pct_change, 1)) %>%
      gt() %>%
      cols_label(vA = input$cmp_period_a, vB = input$cmp_period_b, pct_change = "% Change") %>%
      fmt_number(columns = c(vA, vB), decimals = 0) %>%
      fmt_number(columns = pct_change, decimals = 1) %>%
      tab_header(title = paste0(input$cmp_metric, " — Competitors"))
  })
}

shinyApp(ui, server)
