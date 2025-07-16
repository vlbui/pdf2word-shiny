# 📄 PDF to Word Converter (Shiny App)

Convert any PDF file to a Word document using high-quality image rendering that preserves the original layout.\
This Shiny app offers a visual preview of the first page, page count display, and controls to adjust DPI, shrink factor, and margin trimming.

🎓 **Designed to support graduate students**, especially those preparing a *thesis with publication* as outlined by [Monash University](https://www.monash.edu/graduate-research/support-and-resources/examination/publication).

------------------------------------------------------------------------

## ✨ Features

-   🖼 **Preview the first page** of the PDF
-   🔢 **Display total number of pages**
-   ⚙️ **Custom settings**:
    -   DPI (resolution)
    -   Shrink factor (image scale)
    -   Crop margin (automatic whitespace trimming)
-   📂 **Select output folder**
-   📄 **Convert to Word document (.docx)** with one click
-   🧩 **Automatic package installation**
-   ✅ No admin privileges required

------------------------------------------------------------------------

## 📦 Requirements

-   **R ≥ 4.0**
-   **Internet access (first run only)** to install required packages
-   **Operating systems**: Windows / macOS / Linux

### R Packages (installed automatically):

-   `shiny`
-   `pdftools`
-   `magick`
-   `rmarkdown`
-   `shinyFiles`
-   `fs`

------------------------------------------------------------------------

## 🚀 How to Run

### 🖥 Option 1: Clone and Run in RStudio

``` bash
git clone https://github.com/vlbui/pdf2word-shiny.git
cd pdf2word-shiny
```

Then in RStudio:

``` r
shiny::runApp()
```

Or open `app.R` and click **Run App**.

------------------------------------------------------------------------

### 📦 Option 2: Download ZIP

-   [Download ZIP](https://github.com/vlbui/pdf2word-shiny/archive/refs/heads/main.zip)
-   Unzip the folder
-   Open `app.R` in RStudio
-   Click **Run App**

------------------------------------------------------------------------

## 🧪 How It Works

-   Uses `pdftools` to render each page as a bitmap image (PNG)
-   Automatically crops whitespace using `magick::image_trim()`
-   Rescales the image based on shrink factor
-   Compiles the images into a `.docx` file using `rmarkdown::render()`

📌 This app **does not perform OCR** --- it captures the layout visually.

------------------------------------------------------------------------

## 🧑‍💻 Author

-   **Name**: Viet Bui\
-   **Email**: [viet.bui1\@monash.edu](mailto:viet.bui1@monash.edu), [longbui189\@gmail.com](mailto:longbui189@gmail.com)\
-   **GitHub**: [github.com/vlbui](https://github.com/vlbui)

------------------------------------------------------------------------

## ⚖️ License

This project is licensed under the **MIT License**.\
You are free to use, modify, and distribute it.

------------------------------------------------------------------------

## 🙌 Acknowledgements

Built to help students, researchers, and academics generate clean Word documents from published PDF papers, particularly in support of the *thesis with publication* submission pathway.

> *"Save time formatting --- focus on your research."*
