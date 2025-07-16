# ---- Safe Installer ----
safe_install_packages <- function(pkgs, version_requirements = list()) {
  message("📦 Checking required packages...")
  total <- length(pkgs)
  pb <- txtProgressBar(min = 0, max = total, style = 3)
  for (i in seq_along(pkgs)) {
    pkg <- pkgs[i]
    need_install <- FALSE
    if (!requireNamespace(pkg, quietly = TRUE)) {
      need_install <- TRUE
    } else if (!is.null(version_requirements[[pkg]])) {
      local_version <- as.character(packageVersion(pkg))
      required_version <- version_requirements[[pkg]]
      if (package_version(local_version) < package_version(required_version)) {
        need_install <- TRUE
      }
    }
    if (need_install) {
      tryCatch({
        install.packages(pkg, repos = "https://cloud.r-project.org")
      }, error = function(e) {
        warning(sprintf("❌ Failed to install %s: %s", pkg, e$message))
        showNotification(paste("❌ Failed to install", pkg, ":", e$message),
                         type = "error", duration = NULL)
      })
    }
    setTxtProgressBar(pb, i)
  }
  close(pb)
}

safe_install_packages(
  pkgs = c("shiny", "pdftools", "magick", "rmarkdown", "shinyFiles", "fs"),
  version_requirements = list(magick = "2.7.0")
)

# ---- Load Packages ----
library(shiny)
library(pdftools)
library(magick)
library(rmarkdown)
library(shinyFiles)
library(fs)

# ---- UI ----
ui <- fluidPage(
  titlePanel("📄 PDF to Word (Preview First Page)"),
  
  sidebarLayout(
    sidebarPanel(
      fileInput("pdf", "Choose PDF File", accept = ".pdf"),
      
      shinyDirButton("out_dir", "Select Output Folder", "Choose directory"),
      verbatimTextOutput("chosen_dir"),
      
      sliderInput("dpi", "DPI (Resolution)", 100, 1200, 400, step = 50),
      sliderInput("shrink", "Shrink Factor", 0.1, 1.5, 0.9, step = 0.05),
      sliderInput("crop_margin", "Crop Margin Tolerance (%)", 0, 10, 1, step = 0.5),
      
      actionButton("convert", "🔄 Convert to Word")
    ),
    
    mainPanel(
      h4("📖 First Page Preview + Page Count"),
      uiOutput("pdf_preview"),
      tags$hr(),
      h4("⏱ Timer:"),
      textOutput("timer"),
      tags$hr(),
      h4("📝 Status:"),
      verbatimTextOutput("status")
    )
  ),
  
  tags$hr(),
  tags$footer(
    style = "text-align:center; font-size: 90%; padding: 10px; color: #666;",
    HTML("&copy; 2025 Viet Bui — <a href='mailto:viet.bui1@monash.edu'>viet.bui1@monash.edu</a> &nbsp; | &nbsp; <a href='https://github.com/vlbui/' target='_blank'>GitHub: @vlbui</a><br> This tool is free to use.")
  )
)

# ---- SERVER ----
server <- function(input, output, session) {
  volumes <- c(Home = fs::path_home(), "R Working Dir" = getwd())
  shinyDirChoose(input, "out_dir", roots = volumes, session = session)
  out_dir_path <- reactiveVal(getwd())
  
  observeEvent(input$out_dir, {
    path <- parseDirPath(volumes, input$out_dir)
    out_dir_path(normalizePath(path))
  })
  
  output$chosen_dir <- renderText({
    paste("Output folder:", out_dir_path())
  })
  
  output$pdf_preview <- renderUI({
    req(input$pdf)
    pdf_info <- pdftools::pdf_info(input$pdf$datapath)
    total_pages <- pdf_info$pages
    img_raw <- pdftools::pdf_render_page(input$pdf$datapath, page = 1, dpi = 100)
    img <- magick::image_read(img_raw)
    img_path <- tempfile(fileext = ".png")
    magick::image_write(img, img_path)
    
    tagList(
      tags$p(paste("📄 Total pages:", total_pages)),
      tags$img(src = knitr::image_uri(img_path), style = "max-width: 100%; border: 1px solid #ccc;")
    )
  })
  
  observeEvent(input$convert, {
    req(input$pdf)
    
    start_time <- Sys.time()
    output$status <- renderText("⏳ Converting...")
    output$timer <- renderText("...")
    
    withProgress(message = "Working...", value = 0, {
      incProgress(0.1)
      
      pdf_path <- input$pdf$datapath
      pdf_name <- tools::file_path_sans_ext(input$pdf$name)
      out_dir <- out_dir_path()
      output_docx <- file.path(out_dir, paste0(pdf_name, "_converted.docx"))
      dpi <- input$dpi
      shrink <- input$shrink
      crop_margin_pct <- input$crop_margin / 100
      
      tryCatch({
        tmp_dir <- tempfile("pdf2docx_")
        dir.create(tmp_dir)
        n_pages <- pdftools::pdf_info(pdf_path)$pages
        img_paths <- c()
        
        for (i in seq_len(n_pages)) {
          incProgress(0.5 / n_pages)
          img <- pdf_render_page(pdf_path, page = i, dpi = dpi)
          img <- image_read(img)
          img <- image_trim(img, fuzz = crop_margin_pct * 100)
          info <- image_info(img)
          img <- image_scale(img, paste0(round(info$width * shrink), "\n"))
          img_path <- file.path(tmp_dir, sprintf("page_%03d.png", i))
          image_write(img, img_path)
          img_paths <- c(img_paths, normalizePath(img_path))
        }
        
        rmd_path <- file.path(tmp_dir, "doc.Rmd")
        rmd <- c(
          "---",
          "output: word_document",
          "---",
          "",
          paste0("![](", img_paths, ")")
        )
        writeLines(rmd, rmd_path)
        
        setwd(tmp_dir)
        rmarkdown::render("doc.Rmd", output_file = "rendered.docx", quiet = TRUE)
        file.copy("rendered.docx", output_docx, overwrite = TRUE)
        
        duration <- round(difftime(Sys.time(), start_time, units = "secs"), 1)
        output$timer <- renderText(paste("⏱ Done in", duration, "seconds"))
        output$status <- renderText(paste("✅ Saved to:\n", output_docx))
        
        showModal(modalDialog(
          title = "✅ Conversion Complete",
          paste("Saved to:", output_docx),
          easyClose = TRUE
        ))
      }, error = function(e) {
        output$status <- renderText(paste("❌ Error:", e$message))
        showModal(modalDialog(
          title = "❌ Conversion Failed",
          e$message,
          easyClose = TRUE
        ))
      })
      
      incProgress(1)
    })
  })
}

# ---- RUN APP ----
shinyApp(ui, server)
