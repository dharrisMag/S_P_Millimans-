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

file_path <- "Milliman's Financial Insight v4 (2).xlsx"
wide_raw <- read_excel(file_path, sheet = "Part 2", col_names = FALSE, col_types = "text")

metrics  <- wide_raw[2, -1] %>% unlist(use.names = FALSE) %>% str_squish()
quarters <- wide_raw[4, -1] %>% unlist(use.names = FALSE) %>% str_squish()
col_names <- c("Entity", make.unique(paste0(metrics, "—", quarters)))

data_block <- wide_raw[-(1:4), ]
names(data_block) <- col_names

df <- data_block %>%
  mutate(Entity = str_squish(Entity)) %>%
  pivot_longer(-Entity, names_to = "Metric_Quarter", values_to = "Value") %>%
  separate(Metric_Quarter, into = c("Metric", "Quarter"), sep = "—", extra = "merge") %>%
  mutate(
    Metric  = str_to_upper(str_trim(Metric)),
    Quarter = str_trim(Quarter),
    Value   = as.numeric(str_replace_all(Value, ",", ""))
  ) %>%
  filter(!is.na(Value))

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

all_entities    <- sort(unique(df$Entity))
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

quarters_avail <- intersect(c("2024Q1","2025Q1"), unique(df$Quarter))

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
                 selectizeInput("metric_cmp_all", "Choose Metric:", metrics_list, options = list(placeholder = 'Search metrics...'))
               ),
               mainPanel(
                 gt_output("table_cmp_all"),
                 br()
               )
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
                 selectInput("metric_comp", "Choose a Metric:", choices = metrics_list)
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
  
  df_wide <- df %>%
    filter(Quarter %in% c("2024Q1","2025Q1")) %>%
    pivot_wider(names_from = Quarter, values_from = Value) %>%
    rename(`2024 Q1` = `2024Q1`, `2025 Q1` = `2025Q1`) %>%
    mutate(
      `Absolute Change` = `2025 Q1` - `2024 Q1`,
      `% Change`        = (`Absolute Change`) / `2024 Q1` * 100
    )
  
  output$table_cmp_all <- render_gt({
    req(input$metric_cmp_all)
    tbl <- df_wide %>%
      filter(Metric == input$metric_cmp_all, Entity %in% company_list) %>%
      arrange(desc(`2025 Q1`)) %>%
      select(Company = Entity, `2024 Q1`, `2025 Q1`, `Absolute Change`, `% Change`)
    gt(tbl) %>%
      fmt_number(vars(`2024 Q1`,`2025 Q1`,`Absolute Change`), decimals = 0) %>%
      fmt_number(vars(`% Change`), decimals = 1, pattern = "{x}%") %>%
      tab_header(title = input$metric_cmp_all)
  })
  
  output$entity_select_ui <- renderUI({
    opts <- if (input$snl_pnc_only) {
      snl_pnc_group
    } else if (input$hide_snl_pnc) {
      setdiff(all_entities, snl_pnc_group)
    } else if (input$mode == "Competitors") {
      competitors
    } else {
      all_entities
    }
    selectizeInput("entities", "Filter Entities:", opts, multiple = TRUE)
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
    req(input$metric_comp)
    df %>%
      filter(Metric == input$metric_comp, Entity %in% competitors) %>%
      pivot_wider(names_from = Quarter, values_from = Value) %>%
      rename(v24 = `2024Q1`, v25 = `2025Q1`) %>%
      mutate(pct_change = (v25 - v24) / v24 * 100)
  })
  
  output$comparison_tbl <- render_gt({
    dd <- comp_df()
    gt(dd %>%
         select(Entity, v24, v25, pct_change) %>%
         rename(`2024 Q1` = v24, `2025 Q1` = v25, `% Change` = pct_change) %>%
         mutate(`% Change` = round(`% Change`, 1))
    ) %>%
      fmt_number(columns = c(`2024 Q1`,`2025 Q1`), decimals = 0) %>%
      fmt_number(columns = `% Change`, decimals = 1) %>%
      tab_header(title = input$metric_comp)
  })
  
  output$comparison_plot <- renderPlotly({
    dd <- comp_df()
    p <- ggplot(dd, aes(
      x = Entity, y = pct_change, fill = Entity,
      text = paste0(
        "Entity: ", Entity,
        "<br>2024 Q1: ", round(v24, 0),
        "<br>2025 Q1: ", round(v25, 0),
        "<br>% Change: ", round(pct_change, 1), "%"
      )
    )) +
      geom_col(show.legend = FALSE) +
      scale_fill_manual(values = entity_colors[competitors]) +
      scale_y_continuous(name = "% Change", expand = c(0,0)) +
      labs(x = NULL, title = paste(input$metric_comp, "% Change")) +
      theme_minimal() +
      theme(axis.text.x = element_text(angle = 45, vjust = 1, hjust = 1))
    
    ggplotly(p, tooltip = "text")
  })
  
}

shinyApp(ui, server)
