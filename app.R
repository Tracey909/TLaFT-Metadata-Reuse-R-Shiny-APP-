#’ Title: From TOC to TLaFT
#' Author: Tracy
#' Date: 2026-06
#' Steps:
#' 0. directory setting 
#' 1. packages installing and loading
#' 2. default configuration
#' 3. functions before UI and server
#' 4. UI 
#' 5. Server 
#' 6. app calling


# ===================== 0. directory setting =====================
personal_lib <- "C:/Users/kjhf454/AppData/Local/R/R_Local_Packages"

dir.create(personal_lib, recursive = TRUE, showWarnings = FALSE)

.libPaths(personal_lib)


# ===================== 1. packages installing and loading=====================
zip_list <- list.files(personal_lib, pattern = "\\.zip$", full.names = TRUE)

for (zip_file in zip_list) {
    install.packages(zip_file, repos = NULL, type = "binary", lib = personal_lib)
}

library(shiny)          # Core framework for building interactive web application
library(shinyjs)       # Advanced front-end control: button enable/disable, panel collapse, dynamic UI feedback
library(readxl)         # Read structured Excel files storing TOC, title and footnote metadata
library(openxlsx)       # Write final standardized title-footnote results to formatted Excel outputs
library(DT)             # Render interactive data tables for metadata preview and verification in UI
library(stringr)        # String cleaning, formatting, splicing and text standardization for annotations
library(stringdist)    # Fuzzy matching between TOC tags and pre-defined title-footnote entries
library(lubridate)     # Timestamp logging for audit trail and metadata version recording
library(rhandsontable)  # Editable table component for in-app manual adjustment of metadata before generation
library(fs)             # for recycle use   
library(dplyr)          # for recycle use
# ===================== 2. default configuration =====================
# main structure of final output
CORE_COLUMNS <- c(
    "SHELL_ID", "OUTPUT_TYPE", "INTEXT", "SECTION_TITLE1", "SECTION_TITLE2",
    "OUTPUT_NUMBER", "TITLE1", "TITLE2", "POPULATION", "OUTFILE_NAME",
    "DEV_PROGRAMMER", "QC_PROGRAMMER", "SOURCE_REFERENCE", "MACRO_NAME",
    "PROGRAMMING_NOTE", "FOOTNOTE1", "FOOTNOTE2", "FOOTNOTE3", "FOOTNOTE4",
    "FOOTNOTE5", "FOOTNOTE6", "FOOTNOTE7", "FOOTNOTE8", "FOOTNOTE9", "FOOTNOTE10"
)

# key matching columns  
COLUMN_MAPPING <- list(
    SHELL_ID = "Shell_ID",          
    OUTPUT_TYPE = "Output_Type",    
    OUTPUT_NUMBER = "Output_Number",
    TITLE1 = "Title",               
    POPULATION = "Analysis_Set"     
)

# ===================== 3. functions before UI and server=====================
# function 1: clean text in excel
flatten_col <- function(x) {
    if(is.list(x) || inherits(x, "data.frame")) {
        paste(unlist(x), collapse = " ") %>% str_trim()
    } else {
        as.character(x) %>% str_trim()
    }
}

# function 2: further clean text in excel
text_preprocess <- function(text_vec) {
    text_vec <- sapply(text_vec, flatten_col)
    text_vec[is.na(text_vec)] <- ""
    text_vec <- str_remove_all(text_vec, "[[:punct:]]")
    text_vec <- str_remove_all(text_vec, "v\\d+|revision|trial|version")
    text_vec <- str_squish(text_vec)
    return(text_vec)
}

# function 3: matching method
cosine_match_optimized <- function(toc_title, toc_shell_id, x_titles, x_shell_ids, x_data, top_n = 8, threshold = 0.7) {
    # 安全处理输入数据
    if(!is.data.frame(x_data)) {
        if(is.vector(x_data)) {
            x_data <- data.frame(value = x_data, stringsAsFactors = FALSE)
        } else {
            x_data <- as.data.frame(x_data, stringsAsFactors = FALSE)
        }
    }
    colnames(x_data) <- make.names(colnames(x_data), unique = TRUE)
    
    # 提前返回空数据框（避免后续无效计算）
    if(is.na(toc_title) || nchar(toc_title) < 2 || toc_title == "") {
        return(x_data[0, , drop = FALSE])
    }
    
    # 安全处理输入向量
    x_titles <- if(is.null(x_titles)) character(0) else sapply(x_titles, flatten_col)
    x_titles[is.na(x_titles)] <- ""
    x_shell_ids <- if(is.null(x_shell_ids)) character(0) else sapply(x_shell_ids, flatten_col)
    x_shell_ids[is.na(x_shell_ids)] <- ""
    
    # 提前返回空数据框
    if(all(x_titles == "") & all(x_shell_ids == "")) {
        return(x_data[0, , drop = FALSE])
    }
    
    toc_clean <- text_preprocess(toc_title)
    x_titles_clean <- text_preprocess(x_titles)
    title_sim <- numeric(length(x_titles_clean))
    
    # 计算标题相似度
    for(i in 1:length(x_titles_clean)) {
        x <- x_titles_clean[i]
        if(nchar(x) < 2) {
            title_sim[i] <- 0
        } else {
            title_sim[i] <- round(1 - stringdist(toc_clean, x, method = "cosine"), 3)
        }
    }
    
    # 计算Shell_ID相似度
    shell_id_sim <- ifelse(tolower(x_shell_ids) == tolower(toc_shell_id), 1.0, 0.0)
    avg_sim <- round((title_sim + shell_id_sim) / 2, 3)
    
    # 核心修复：先判断是否有符合阈值的结果，再进行后续操作
    if(all(avg_sim < threshold)) {
        return(x_data[0, , drop = FALSE])
    }
    
    # 只处理符合阈值的数据（避免空数据框赋值）
    x_data_temp <- x_data[avg_sim >= threshold, , drop = FALSE]
    if(nrow(x_data_temp) == 0) {
        return(x_data[0, , drop = FALSE])
    }
    
    # 给筛选后的数据添加相似度列（此时x_data_temp一定有行）
    x_data_temp$avg_similarity <- avg_sim[avg_sim >= threshold]
    x_data_temp$title_similarity <- title_sim[avg_sim >= threshold]
    x_data_temp$shell_id_similarity <- shell_id_sim[avg_sim >= threshold]
    
    # 处理日期排序（仅当有数据时）
    if("DATE" %in% colnames(x_data_temp)) {
        DATE_parsed <- rep(as.Date("1970-01-01"), nrow(x_data_temp))
        for(i in 1:nrow(x_data_temp)) {
            dt <- x_data_temp$DATE[i]
            DATE_parsed[i] <- tryCatch({
                as.Date(dt, format = "%Y-%m-%d")
            }, error = function(e) {
                tryCatch({
                    as.Date(dt, format = "%Y/%m/%d")
                }, error = function(e) {
                    tryCatch({
                        as.Date(dt, format = "%m/%d/%Y")
                    }, error = function(e) {
                        as.Date("1970-01-01")
                    })
                })
            })
        }
        x_data_temp$DATE_parsed <- DATE_parsed
        ord <- order(-x_data_temp$avg_similarity, -as.numeric(x_data_temp$DATE_parsed))
        x_data_temp <- x_data_temp[ord, , drop = FALSE]
        x_data_temp <- x_data_temp[, setdiff(colnames(x_data_temp), "DATE_parsed"), drop = FALSE]
    } else {
        ord <- order(-x_data_temp$avg_similarity)
        x_data_temp <- x_data_temp[ord, , drop = FALSE]
    }
    
    # 取TOP N结果
    top_n <- min(top_n, nrow(x_data_temp))
    match_result <- x_data_temp[1:top_n, , drop = FALSE]
    
    # 移除相似度列，返回原始列
    match_result <- match_result[, setdiff(colnames(match_result), 
                                           c("avg_similarity", "title_similarity", "shell_id_similarity")), 
                                 drop = FALSE]
    
    return(as.data.frame(match_result, stringsAsFactors = FALSE))
}

# function 4: tailor TOC column into TLaFT columns
split_output_number <- function(output_num_str) {
    output_num_str <- flatten_col(output_num_str)
    output_num_str[is.na(output_num_str)] <- ""
    output_num_str <- str_trim(output_num_str)
    
    if(output_num_str == "") {
        return(list(OUTPUT_TYPE = "", OUTPUT_NUMBER = ""))
    }
    
    clean_for_type <- str_remove_all(output_num_str, "[[:space:][:punct:]]")
    first_char <- if(nchar(clean_for_type) > 0) substr(clean_for_type, 1, 1) else ""
    OUTPUT_TYPE <- str_to_upper(first_char)
    
    num_parts <- str_extract_all(output_num_str, "[0-9.]+")[[1]]
    OUTPUT_NUMBER <- paste(num_parts, collapse = "")
    OUTPUT_NUMBER <- str_remove_all(OUTPUT_NUMBER, "^\\.+|\\.+$")
    OUTPUT_NUMBER <- ifelse(OUTPUT_NUMBER == "", "", OUTPUT_NUMBER)
    
    return(list(OUTPUT_TYPE = OUTPUT_TYPE, OUTPUT_NUMBER = OUTPUT_NUMBER))
}

# function 5: execute functions in cart with filtered records
process_cart_data <- function(x_match_row, toc_data) {
    # 安全检查：确保输入是数据框且有列名
    if(!is.data.frame(x_match_row)) {
        x_match_row <- data.frame()
    }
    if(!is.data.frame(toc_data)) {
        toc_data <- data.frame()
    }
    
    # 步骤1：统一列名格式（空数据框处理）
    if(ncol(x_match_row) > 0) {
        colnames(x_match_row) <- toupper(make.names(colnames(x_match_row), unique = TRUE))
    }
    if(ncol(toc_data) > 0) {
        colnames(toc_data) <- toupper(make.names(colnames(toc_data), unique = TRUE))
    }
    
    # 步骤2：初始化空数据框（修复核心：确保数据框结构稳定）
    cart_matrix <- matrix("", nrow = 1, ncol = length(CORE_COLUMNS))
    cart_data <- as.data.frame(cart_matrix, stringsAsFactors = FALSE)
    colnames(cart_data) <- CORE_COLUMNS
    
    # 步骤3：填充X中的非空值（添加长度判断，避免if参数长度为零）
    if(nrow(x_match_row) > 0 && ncol(x_match_row) > 0) {
        for(col in CORE_COLUMNS) {
            x_col <- grep(paste0("^", col, "$"), colnames(x_match_row), ignore.case = TRUE, value = TRUE)
            # 核心修复：先判断x_col是否有值（长度>0）
            if(length(x_col) > 0) {
                x_val <- flatten_col(x_match_row[[x_col]])
                # 再判断值是否非空
                if(nchar(x_val) > 0 && x_val != "") {
                    cart_data[[col]] <- x_val
                }
            }
        }
    }
    
    # 步骤4：替换为TOC的对应值（核心规则，添加长度判断）
    if(nrow(toc_data) > 0 && ncol(toc_data) > 0) {
        for(x_col in names(COLUMN_MAPPING)) {
            toc_col <- COLUMN_MAPPING[[x_col]]
            toc_col_match <- grep(paste0("^", toc_col, "$"), colnames(toc_data), ignore.case = TRUE, value = TRUE)
            # 核心修复：先判断匹配结果是否有值
            if(length(toc_col_match) > 0) {
                toc_val <- flatten_col(toc_data[[toc_col_match]])
                # 再判断值是否非空
                if(nchar(toc_val) > 0 && toc_val != "") {
                    cart_data[[x_col]] <- toc_val
                }
            }
        }
        
        # 特殊处理：Output_Number拆分后替换（添加长度判断）
        toc_output_num_col <- grep("^OUTPUT_NUMBER$", colnames(toc_data), ignore.case = TRUE, value = TRUE)
        if(length(toc_output_num_col) > 0) {
            split_res <- split_output_number(toc_data[[toc_output_num_col]])
            # 修复：判断split_res返回值是否非空
            if(nchar(split_res$OUTPUT_TYPE) > 0 && split_res$OUTPUT_TYPE != "") {
                cart_data[["OUTPUT_TYPE"]] <- split_res$OUTPUT_TYPE
            }
            if(nchar(split_res$OUTPUT_NUMBER) > 0 && split_res$OUTPUT_NUMBER != "") {
                cart_data[["OUTPUT_NUMBER"]] <- split_res$OUTPUT_NUMBER
            }
        }
    }
    
    # 添加TOC_No（添加长度判断）
    toc_no_val <- ""
    if(nrow(toc_data) > 0 && ncol(toc_data) > 0) {
        toc_no_col <- grep("^NO$", colnames(toc_data), ignore.case = TRUE, value = TRUE)
        if(length(toc_no_col) > 0) {
            toc_no_val <- as.character(flatten_col(toc_data[[toc_no_col]]))
        }
    }
    cart_data$TOC_No <- ifelse(nchar(toc_no_val) > 0, toc_no_val, "")
    
    # 调整列序
    col_order <- c("TOC_No", CORE_COLUMNS)
    cart_data <- cart_data[, col_order, drop = FALSE]
    
    # 兜底：确保返回有效数据框
    if(!is.data.frame(cart_data) || nrow(cart_data) == 0) {
        cart_matrix_empty <- matrix("", nrow = 1, ncol = length(CORE_COLUMNS))
        cart_data_empty <- as.data.frame(cart_matrix_empty, stringsAsFactors = FALSE)
        colnames(cart_data_empty) <- CORE_COLUMNS
        cart_data_empty$TOC_No <- toc_no_val
        cart_data_empty <- cart_data_empty[, col_order, drop = FALSE]
        cart_data <- cart_data_empty
    }
    
    return(cart_data)
}

# ===================== 4. UI =====================

ui <- fluidPage(
    useShinyjs(),
    #ui 1: title panel setting
tags$head(tags$meta(charset = "UTF-8")),
titlePanel(
  tags$span(
    "TOC to TLaFT",
    style = "
    font-size:24px !important; 
    font-weight:bold; 
    background: linear-gradient(180deg, #212121 0%, #616161 100%);
    -webkit-background-clip: text;
    color: transparent;
    margin:0; 
    padding:0; 
    letter-spacing:0.3px;
    "
  )
)    ,
    #ui 2: main container 
    tags$style(HTML("
    .main-container { 
    display: flex; height: calc(100vh - 80px); position: relative; }
    
     /**** left panel 1 ****/
    .input-panel { 
      background-color: #f8f9fa; 
      padding: 4px !important; 
      border-right:1px solid #dee2e6; 
      font-size: 12px !important;  
      width: 33.333%; 
      min-width: 200px;
      transition: all 0.3s ease; 
      box-sizing: border-box;  
      display: flex; 
      flex-direction: column; 
      height: 100%;
    }
    
     /**** items in left panel TLaFT upload****/
    .upload-row {
      display: flex;
      flex-direction: row; 
      gap: 4px !important; 
      margin-bottom: 6px !important; 
      flex-shrink: 0; 
    }
    
    .upload-item {
      flex: 1; 
      min-width: 150px; 
    }
    
    .upload-item.loaded {
      background-color: #e9ecef !important;
      color: #212529 !important;
      border-radius: 4px;
      padding: 4px;
      transition: background-color 0.3s ease;
    }
    .upload-item.loaded select {
      color: #212529 !important;
      background-color: #f8f9fa !important;
      border-color: #dee2e6 !important;
    }
    
      /**** items in left panel TOC upload****/
   
    .toc-select-btn {
      padding: 2px 2px; 
      flex-shrink: 0;
      display: flex;
      align-items: center;
      justify-content: flex-start;
    }
    .toc-select-btn .btn-xs {
      font-size: 9px !important;
      padding: 1px 4px !important;
      height: 18px !important;
      width: 24px !important;
      border: 1px solid #dee2e6 !important;
      display: flex;
      align-items: center;
      justify-content: center;
    }
    
      /**** items in left panel TOC display****/
    
    .toc-table-wrapper {
      flex: 1;
      margin-top: 0 !important;
      border: 1px solid #eee; 
      min-height: 0;
    }
    .toc-table-wrapper .dataTables_wrapper {
      width: 100% !important;
      font-size: 9px;
    }
    .toc-table-wrapper .dataTables_scroll {
      overflow: auto !important;
    }
    .toc-table-wrapper .dataTables_scrollHead {
      overflow: hidden;
    }
    .toc-table-wrapper .dataTables_scrollBody {
      overflow-y: auto !important;
      overflow-x: auto !important;
      border:0 !important;
    }
    
      /**** items in left panel fileinput****/
    
    .input-panel .file-input {
       width: 100% !important; 
       height: 20px !important;
       margin-top: 4px !important;
     }
    .input-panel .file-input .btn {
      width: 100px !important;  
      font-size: 12px !important; 
      padding: 4px 8px !important;  
      height: 20px !important 
    }
    .input-panel .file-input .form-control {
      font-size: 12px !important;  
      height: 20px !important 
    }
    .input-panel .card {
      padding: 6px !important; 
      margin-bottom: 0 !important; 
      border-radius: 4px !important; 
    }
    
      /**** items in right panel ****/
    
    .output-panel { 
      flex: 1; padding:4px; overflow-y:hidden !important; overflow-x:hidden !important; transition: all 0.3s ease;
    }
    
    .cart-highlight { background-color: #e7f5ff !important; font-weight: 500; }
    .sheet-select { width: 100% !important; margin: 4px 0 !important; }
    .verbatim-small { font-size:9px !important; }
    .verbatim-status { font-size:9px !important; color:#6c757d !important; }
    .resizer {
      position: absolute; top: 0; bottom: 0; left: 33.333%; width: 4px;
      background: transparent; cursor: ew-resize; z-index: 9998;
    }
    .resizer:hover { background: #dee2e6; border-radius: 2px; }
    .input-panel.hidden { width: 0 !important; padding:0 !important; border:0 !important; }
    .resizer.hidden-indicator { left: 2px !important; background: #6c757d !important; }
    
    /* 新增：上下分隔栏样式 */
    .vertical-resizer {
      position: relative;
      width: 100%;
      height: 4px;
      background: transparent;
      cursor: ns-resize;
      z-index: 9998;
    }
    .vertical-resizer:hover {
      background: #dee2e6;
      border-radius: 2px;
    }
    /* 隐藏Matching Results时的分隔栏样式 */
    .vertical-resizer.hidden-indicator {
      height: 4px;
      background: #6c757d !important;
    }
    /* Matching Results隐藏时的样式 */
    .match-panel.hidden {
      height: 0 !important;
      padding: 0 !important;
      border: 0 !important;
      overflow: hidden;
    }
    /* ========== 核心优化：Cart面板样式 ========== */
    /* 1. Match面板：关闭外层滚动，仅表格滚动 */
    .match-panel {
      flex: 1; 
      min-height: 100px; 
      overflow-y: hidden !important; 
      overflow-x: hidden !important; 
      transition: all 0.3s ease;
      padding: 0 !important;
    }
    .match-panel .card {
      height: 100% !important;
      display: flex;
      flex-direction: column;
      padding: 10px;
    }
    .match-panel .card > div:last-child {
      flex: 1;
      height: 100%;
      overflow-y:  hidden !important; 
      overflow-x:  hidden !important; 
      /* 可选：优化滚动条样式，避免出现双滚动条时的挤压 */
      box-sizing: border-box;
      padding-bottom: 8px; /* 给横向滚动条预留空间 */
    }
    /* 2. Cart+导出面板：隐藏匹配后占满100%，flex垂直布局 */
    .cart-export-panel {
      flex: 1; 
      min-height: 200px; 
      display: flex; 
      flex-direction: column;  
      transition: all 0.3s ease;
      overflow-y: hidden !important; /* 关闭外层滚动，仅表格滚动 */
      overflow-x: hidden !important; /* 关闭外层滚动，仅表格滚动 */
      padding: 0 !important;
    }
    /* 3. Cart面板内的卡片：强制占满父容器100%高度 */
    .cart-export-panel .card {
      height: 100% !important;
      display: flex;
      flex-direction: column;
      margin-bottom: 0 !important;
      padding: 10px;
    }
    /* 4. Cart表格容器：核心！flex:1占满除操作行外的所有高度 */
    .cart-table-container {
      flex: 1 !important;
      overflow-y:  hidden !important; 
      overflow-x:  hidden !important; 
      box-sizing: border-box;
      padding-bottom: 8px; /* 预留横向滚动条空间，避免遮挡表格最后一行 */
    }
    /* 5. 操作行：固定高度，不挤压表格 */
    .export-row {
      display: flex;
      gap: 8px;
      align-items: center;
      width: 100%;
      flex-shrink: 0 !important; /* 关键：固定高度，不随表格压缩 */
      margin: 0 !important;
      padding-top: 8px;
      border-top: 1px solid #eee; /* 加分割线，区分表格和操作行 */
    }
    .rhandsontable.html-widget {
    width: 100% !important;
    /* 关键：取消固定宽度，让表格内容决定是否触发横向滚动 */
    min-width: 100%; /* 确保表格至少占满容器宽度 */
    height: 100% !important; /* 表格占满父容器高度，触发纵向滚动 */
      }
  ")),
    
    tags$div(class = "main-container",
             tags$div(class = "resizer", id = "resizer"),
             
             column(width = 12, class = "input-panel",
                    
                    div(
                        style = "display:flex; gap:4px; align-items:center; height:38px;",
                        
                        tags$input(
                            id = "recycle_target_folder",
                            type = "text",
                            class = "form-control",
                            placeholder = "Please enter target folder path...",
                            style = "flex:1; height:100%; padding:0 10px; box-sizing:border-box;"
                        ),
                        
                        actionButton(
                            inputId = "run_recycle_tlaft",
                            label = "Recycle",
                            class = "btn btn-primary btn-sm",
                            style = "flex-shrink:0; white-space:nowrap; height:100%; padding:0 15px; display:flex; align-items:center; justify-content:center;"
                        )
                    ),
                    
                    div(class = "card upload-row",style = "margin-top: 16px;",
                        div(class = "upload-item", id = "x_file_container",
                            h4("📁 Recycled TLaFT", style="margin-top:0; font-size:14px; margin-bottom:0 !important;"),
                            fileInput("upload_x", "", accept = c(".xlsx", ".xls"), width="100%"),
                            uiOutput("x_sheet_select_ui")
                        ),
                        div(class = "upload-item", id = "toc_file_container",
                            h4("🎯 TOC File", style="margin-top:0; font-size:14px; margin-bottom:0 !important;"),
                            fileInput("upload_toc", "", accept = c(".xlsx", ".xls"), width="100%"),
                            uiOutput("toc_sheet_select_ui")
                        )
                    ),
                    
                    div(class = "card toc-table-wrapper",
                        # 动态渲染选择按钮
                        uiOutput("toc_select_button_ui"),
                        DTOutput("toc_table", width = "100%")
                    )
             ),
             
             # 右侧面板（优化后）
             column(width = 12, class = "output-panel",
                    # 外层垂直容器（承载上下两个区域+分隔栏）
                    tags$div(class = "output-vertical-container", style = "display: flex; flex-direction: column; height: 100%; width: 100%; position: relative;",
                             # 1. Matching Results 区域
                             tags$div(class = "match-panel", style = "flex: 40%; min-height: 100px; transition: all 0.3s ease;",
                                      div(class = "card",
                                          h4("✅ Matching Results", style="margin-top:0; font-size:14px; margin-bottom:8px !important;"),
                                          rHandsontableOutput("match_result_table", width = "100%")                                    )
                             ),
                             # 新增：上下分隔栏
                             tags$div(class = "vertical-resizer", id = "vertical-resizer"),
                             # 2. Cart + 导出操作区域（优化结构）
                             tags$div(class = "cart-export-panel", 
                                      style = "flex: 60%; min-height: 200px; display: flex; flex-direction: column; transition: all 0.3s ease;",
                                      div(class = "card",
                                          style = "display: flex; flex-direction: column; height: 100%;",
                                          h4("🛒 Cart (Editable)", style="margin-top:0; font-size:14px; margin-bottom:8px !important;"),
                                          # Cart表格：专属容器，占满剩余高度
                                          div(class = "cart-table-container",
                                              rHandsontableOutput("cart_table", width = "100%")
                                          ),
                                          # 导出操作行：固定高度
                                          div(class = "export-row", 
                                              actionButton("clear_cart", "Clear Cart", class = "btn btn-danger btn-sm", style = "flex: 0 0 100px;"),
                                              actionButton("export_z", "Export", class = "btn btn-success btn-sm", style = "flex: 0 0 100px;")
                                          ),
                                          # 导出状态提示
                                          verbatimTextOutput("export_status") %>% tagAppendAttributes(class = "verbatim-small")
                                      )
                             )
                    )
             )
    ),
    
    tags$script(HTML("
  let isHidden = false;
  let isDragging = false;
  const resizer = document.getElementById('resizer');
  const inputPanel = document.querySelector('.input-panel');
  const container = document.querySelector('.main-container');
  const tocTableWrapper = document.querySelector('.toc-table-wrapper');
  
  function syncTableScroll() {
    const scrollBody = tocTableWrapper.querySelector('.dataTables_scrollBody');
    const scrollHead = tocTableWrapper.querySelector('.dataTables_scrollHeadInner');
    if(scrollBody && scrollHead) {
      scrollBody.addEventListener('scroll', function() {
        scrollHead.style.transform = `translateX(-${this.scrollLeft}px)`;
      });
      const thElements = tocTableWrapper.querySelectorAll('thead th');
      thElements.forEach(th => {
        th.addEventListener('mousedown', function(e) {
          if(e.offsetX > th.offsetWidth - 10) {
            setTimeout(() => {
              if(window.$ && window.$.fn.dataTable) {
                const table = $(tocTableWrapper).find('table').DataTable();
                table.columns.adjust().draw();
              }
            }, 100);
          }
        });
      });
    }
  }
  
  resizer.addEventListener('mousedown', function(e) {
    if (e.button !== 0) return;
    e.preventDefault();
    isDragging = true;
    const startX = e.pageX;
    const startWidth = inputPanel.getBoundingClientRect().width;
    const containerWidth = container.getBoundingClientRect().width;
    
    function resize(e) {
      if(isHidden || !isDragging) return;
      const width = startWidth + (e.pageX - startX);
      let widthPercent = (width / containerWidth) * 100;
      widthPercent = Math.max(5, Math.min(95, widthPercent));
      inputPanel.style.width = widthPercent + '%';
      resizer.style.left = widthPercent + '%';
      
      if(window.$ && window.$.fn.dataTable) {
        const table = $(tocTableWrapper).find('table').DataTable();
        if(table) table.columns.adjust().draw();
      }
    }
    
    function stopResize() {
      isDragging = false;
      document.removeEventListener('mousemove', resize);
      document.removeEventListener('mouseup', stopResize);
      syncTableScroll();
    }
    
    document.addEventListener('mousemove', resize);
    document.addEventListener('mouseup', stopResize);
  });
  
  resizer.addEventListener('contextmenu', function(e) {
    e.preventDefault();
    if(isDragging) return;
    
    isHidden = !isHidden;
    if(isHidden) {
      inputPanel.classList.add('hidden');
      resizer.style.left = '2px';
      resizer.classList.add('hidden-indicator');
    } else {
      const lastWidth = inputPanel.style.width || '33.333%';
      inputPanel.classList.remove('hidden');
      inputPanel.style.width = lastWidth;
      resizer.style.left = lastWidth;
      resizer.classList.remove('hidden-indicator');
      setTimeout(() => {
        if(window.$ && window.$.fn.dataTable) {
          const table = $(tocTableWrapper).find('table').DataTable();
          if(table) table.columns.adjust().draw();
        }
      }, 100);
    }
  });
  
  window.hideSidebar = function() {
    if(!isHidden && !isDragging) {
      isHidden = true;
      inputPanel.classList.add('hidden');
      resizer.style.left = '2px';
      resizer.classList.add('hidden-indicator');
    }
  };
  
  window.showSidebar = function() {
    if(isHidden && !isDragging) {
      isHidden = false;
      const lastWidth = inputPanel.style.width || '33.333%';
      inputPanel.classList.remove('hidden');
      inputPanel.style.width = lastWidth;
      resizer.style.left = lastWidth;
      resizer.classList.remove('hidden-indicator');
      setTimeout(() => {
        if(window.$ && window.$.fn.dataTable) {
          const table = $(tocTableWrapper).find('table').DataTable();
          if(table) table.columns.adjust().draw();
        }
      }, 100);
    }
  };
  
  // 上下分隔栏逻辑
  let isMatchPanelHidden = false;
  let isVerticalDragging = false;
  const verticalResizer = document.getElementById('vertical-resizer');
  const matchPanel = document.querySelector('.match-panel');
  const cartExportPanel = document.querySelector('.cart-export-panel');
  const outputVerticalContainer = document.querySelector('.output-vertical-container');
  
  // 上下拖拽调整高度
  verticalResizer.addEventListener('mousedown', function(e) {
    if (e.button !== 0) return;
    e.preventDefault();
    isVerticalDragging = true;
    const startY = e.pageY;
    const containerHeight = outputVerticalContainer.getBoundingClientRect().height;
    const startMatchHeight = matchPanel.getBoundingClientRect().height;
    
    function resize(e) {
      if(isMatchPanelHidden || !isVerticalDragging) return;
      const deltaY = e.pageY - startY;
      let newMatchHeight = startMatchHeight + deltaY;
      // 限制最小高度
      newMatchHeight = Math.max(100, Math.min(containerHeight - 200, newMatchHeight));
      const matchPercent = (newMatchHeight / containerHeight) * 100;
      const cartPercent = 100 - matchPercent;
      
      matchPanel.style.flex = `${matchPercent}%`;
      cartExportPanel.style.flex = `${cartPercent}%`;
    }
    
    function stopResize() {
      isVerticalDragging = false;
      document.removeEventListener('mousemove', resize);
      document.removeEventListener('mouseup', stopResize);
      // 强制刷新Cart表格高度
      if(window.Shiny) {
        Shiny.setInputValue('refresh_cart_table', Math.random());
      }
    }
    
    document.addEventListener('mousemove', resize);
    document.addEventListener('mouseup', stopResize);
  });
  
  // 右键隐藏/显示Matching Results（优化版）
  verticalResizer.addEventListener('contextmenu', function(e) {
    e.preventDefault();
    if(isVerticalDragging) return;
    
    isMatchPanelHidden = !isMatchPanelHidden;
    if(isMatchPanelHidden) {
      matchPanel.classList.add('hidden');
      verticalResizer.classList.add('hidden-indicator');
      cartExportPanel.style.flex = '100%'; // 占满整个右侧面板
      // 强制重绘，保证flex布局生效
      cartExportPanel.style.display = 'none';
      setTimeout(() => {cartExportPanel.style.display = 'flex'}, 50);
    } else {
      matchPanel.classList.remove('hidden');
      verticalResizer.classList.remove('hidden-indicator');
      matchPanel.style.flex = '40%'; 
      cartExportPanel.style.flex = '60%';
    }
    // 强制刷新表格高度
    if(window.Shiny) {
      Shiny.setInputValue('refresh_cart_table', Math.random());
    }
  });
  
  // 辅助函数
  window.hideMatchPanel = function() {
    if(!isMatchPanelHidden && !isVerticalDragging) {
      isMatchPanelHidden = true;
      matchPanel.classList.add('hidden');
      verticalResizer.classList.add('hidden-indicator');
      cartExportPanel.style.flex = '100%';
      cartExportPanel.style.display = 'none';
      setTimeout(() => {cartExportPanel.style.display = 'flex'}, 50);
    }
  };
  
  window.showMatchPanel = function() {
    if(isMatchPanelHidden && !isVerticalDragging) {
      isMatchPanelHidden = false;
      matchPanel.classList.remove('hidden');
      verticalResizer.classList.remove('hidden-indicator');
      matchPanel.style.flex = '40%';
      cartExportPanel.style.flex = '60%';
    }
  };
  
  document.addEventListener('DOMContentLoaded', function() {
    setTimeout(syncTableScroll, 500);
  });
  
  $(document).on('shiny:value', function(e) {
    if(e.name === 'toc_table') {
      setTimeout(syncTableScroll, 100);
    }
  });
  
$(document).on('afterRender', '.match-panel .hot-container', function(){
  const hot = $(this).handsontable('getInstance');
  // 按下鼠标立刻拦截原生单元格点击，优先执行整行选中，适配只读单元格
  hot.addHook('beforeCellMouseDown', function(event, coords){
    event.preventDefault();
    hot.selectRows(coords.row, coords.row);
    return false; // 阻断原生只读单元格默认点击行为，避免选中延迟
  });
  // 强制锁定只能单行整行选中，禁止单元格局部选区
  hot.addHook('afterSelection', function(row, col, row2, col2){
    if(row !== row2 || col !== col2){
      this.selectRows(row, row);
    }
  });
  // 右键唤起菜单前预选中，双重兜底
  hot.addHook('beforeContextMenuShow', function() {
    const selected = this.getSelected();
    if (!selected || selected.length === 0) {
      const hover = this.getCellCoordsFromMouse(event);
      if(hover.row >= 0) this.selectRows(hover.row);
    }
  })
}) ;    
")))

# ===================== 5. Server =====================


server <- function(input, output, session) {
####Recycle TLaFT firstly####    
    ####Recycle TLaFT firstly####    
    # ===================== Recycle TLaFT 内置工具函数（优化加固） =====================
    unify_col_type <- function(df, target_cols = NULL) {
        if (is.null(target_cols)) {
            df <- df %>% mutate(across(everything(), as.character))
        } else {
            for (col in target_cols) {
                if (col %in% colnames(df)) {
                    df[[col]] <- as.character(df[[col]])
                }
            }
        }
        return(df)
    }
    
    check_writable <- function(dir) {
        if (!dir.exists(dir)) return(FALSE)
        test_file <- tempfile(tmpdir = dir)
        writable <- tryCatch({
            file.create(test_file); file.remove(test_file); TRUE
        }, error = function(e) FALSE)
        return(writable)
    }
    
    # 新增：安全提取相对路径，兜底填充Unknown
    extract_relative_path <- function(full_path) {
        raw_rel <- str_extract(full_path, "(?<=E:/ar-dev/).*(?=/metadata/system)")
        if (is.na(raw_rel) || raw_rel == "") return("Unknown_Project_Path")
        raw_rel
    }
    
    # 新增：通用样式应用函数（CN/EN双sheet复用，避免代码重复）
    apply_sheet_style <- function(wb, sheet_name, sheet_data) {
        # 冻结窗格（与原逻辑一致：首行冻结、前3列冻结）
        freezePane(wb, sheet = sheet_name, firstActiveRow = 2, firstActiveCol = 4)
        
        # 样式定义
        table_style <- createStyle(
            fontName = "Calibri", fontSize = 12, fgFill = "#F2F2F2",
            halign = "center", valign = "center", textDecoration = "bold",
            border = "TopBottomLeftRight", borderColour = "#BFBFBF"
        )
        content_style <- createStyle(
            fontName = "Calibri", fontSize = 10, halign = "left", valign = "center",
            border = "TopBottomLeftRight", borderColour = "#BFBFBF"
        )
        zebra_style <- createStyle(fgFill = "#F8F9FA", border = "TopBottomLeftRight", borderColour = "#BFBFBF")
        full_border <- createStyle(border = "TopBottomLeftRight", borderColour = "#BFBFBF")
        
        # 写入数据
        writeData(wb, sheet = sheet_name, x = sheet_data, withFilter = TRUE, headerStyle = table_style)
        
        # 应用内容样式、斑马条纹、全边框
        row_count <- nrow(sheet_data)
        col_count <- ncol(sheet_data)
        if (row_count > 0) {
            addStyle(wb, sheet_name, content_style, rows = 2:(row_count+1), cols = 1:col_count, gridExpand = TRUE, stack = TRUE)
            addStyle(wb, sheet_name, zebra_style, rows = seq(3, row_count+1, 2), cols = 1:col_count, gridExpand = TRUE, stack = TRUE)
        }
        addStyle(wb, sheet_name, full_border, rows = 1:(row_count+1), cols = 1:col_count, gridExpand = TRUE, stack = TRUE)
        
        # 自适应列宽（与原逻辑一致，空内容兜底）
        setColWidths(wb, sheet_name, cols = 1:col_count,
                     widths = sapply(1:col_count, function(col_idx) {
                         col_name <- colnames(sheet_data)[col_idx]
                         col_content <- sheet_data[[col_idx]]
                         header_len <- nchar(col_name)
                         content_len <- if(all(is.na(col_content)) || all(trimws(as.character(col_content))=="")) 0 else max(nchar(as.character(col_content)), na.rm = TRUE)
                         max_len <- max(header_len, content_len)
                         auto_width <- max_len * 1.1 + 1
                         min(max(auto_width, header_len), 60)
                     }))
    }
    
    # ===================== 回收合并主触发逻辑（点击按钮执行） =====================
    observeEvent(input$run_recycle_tlaft, {
        # 全局单次编码设置，退出后恢复原始编码，不污染全局
        old_enc <- getOption("encoding")
        on.exit(options(encoding = old_enc))
        options(encoding = "UTF-8")
        
        # 前置校验：路径非空
        target_folder <- trimws(input$recycle_target_folder)
        if (target_folder == "") {
            showNotification("❌ Please input Target Folder Path first", type = "error")
            output$recycle_status <- renderPrint({cat("Error: Empty folder path")})
            return()
        }
        if (!dir.exists(target_folder)) {
            showNotification("❌ Target folder does not exist", type = "error")
            output$recycle_status <- renderPrint({cat(paste("Invalid Path:", target_folder))})
            return()
        }
        
        # 业务参数集中配置
        default_output_folder <- getwd()
        keyword <- "title-footnote"
        output_file <- "Recycled_TLaFT.xlsx"
        
        # 进度条提升体验，适配批量文件读取耗时场景
        withProgress(message = "Recycle TLaFT...", value = 0, detail = "Scanning Excel files", {
            tryCatch({
                # 1. 遍历目标Excel，过滤临时锁定文件
                excel_files <- dir_ls(
                    path = target_folder,
                    glob = "*.xlsx",
                    recurse = TRUE,
                    ignore.case = TRUE,
                    fail = FALSE
                ) %>% 
                    str_subset(pattern = keyword) %>%
                    { .[!str_starts(path_file(.), "~$")] }
                
                if (length(excel_files) == 0) {
                    stop(paste0("No matched Excel with keyword: ", keyword))
                }
                incProgress(0.1, detail = paste("Detected", length(excel_files), "target files"))
                
                # 局部初始化，杜绝全局变量累积
                success_files <- character(0)
                failed_files <- character(0)
                all_data_list <- list()
                total_file_cnt <- length(excel_files)
                
                # 2. 循环读取所有匹配Excel
                for (idx in seq_along(excel_files)) {
                    file_path <- excel_files[idx]
                    file_name <- path_file(file_path)
                    # 修复进度条上限，最终可达到0.9
                    progress_pct <- 0.1 + 0.8 * (idx / total_file_cnt)
                    setProgress(value = progress_pct, detail = paste("Reading:", file_name))
                    
                    # 二次兜底跳过临时文件
                    if (str_starts(file_name, "~$")) {
                        msg <- paste0(file_name, " (Excel Temp File Skipped)")
                        failed_files <- c(failed_files, msg)
                        next
                    }
                    if (!file.exists(file_path)) {
                        msg <- paste0(file_name, " (Not Found)")
                        failed_files <- c(failed_files, msg)
                        next
                    }
                    if (file.access(file_path, mode = 4) != 0) {
                        msg <- paste0(file_name, " (Permission Denied)")
                        failed_files <- c(failed_files, msg)
                        next
                    }
                    
                    # ====== 关键容错：工作表不存在则跳过文件，不中断整体流程 ======
                    single_df <- tryCatch({
                        read_excel(
                            path = file_path,
                            sheet = 'Title & Footnote',
                            na = "",
                            trim_ws = TRUE
                        )
                    }, error = function(e) {
                        err_txt <- tolower(e$message)
                        if (grepl("sheet.*not found", err_txt)) {
                            msg <- paste0(file_name, " → Missing sheet: Title & Footnote, skipped")
                            failed_files <<- c(failed_files, msg)
                            return(NULL)
                        } else {
                            msg <- paste0(file_name, " → Read failed: ", e$message)
                            failed_files <<- c(failed_files, msg)
                            return(NULL)
                        }
                    })
                    
                    # 读取失败/工作表缺失，直接跳过当前文件
                    if (is.null(single_df)) next
                    
                    # Output_Number格式修正逻辑完全保留
                    num_col <- grep("output_number", colnames(single_df), ignore.case = TRUE)
                    if (length(num_col) > 0) {
                        single_df[[num_col]] <- sapply(single_df[[num_col]], function(x) {
                            if (is.na(x) || x == "" || trimws(x) == "") return(x)
                            dot_count <- str_count(x, "\\.")
                            if (dot_count == 1) {
                                return(str_remove(trimws(sprintf("%.15f", as.numeric(x))), "\\.?0+$"))
                            } else {
                                return(x)
                            }
                        })
                    }
                    
                    single_df <- unify_col_type(single_df)
                    df_with_info <- single_df %>%
                        mutate(
                            Path = extract_relative_path(file_path),
                            File = file_name,
                            Date = substr(as.character(file.info(file_path)$mtime), 1, 10)
                            
    
                        ) %>% relocate(c(Path, Date, File), .before = 1)
                    
                    all_data_list[[file_name]] <- df_with_info
                    success_files <- c(success_files, file_name)
                }
                
                incProgress(0.9, detail = "Combining datasets & splitting CN/EN sheets")
                # 3. 合并全量数据、统一列名大小写
                if (length(all_data_list) == 0) stop("No valid readable data")
                unified_list <- lapply(all_data_list, function(df){
                    colnames(df) <- tolower(colnames(df))
                    df
                })
                total_data <- bind_rows(unified_list) %>% arrange(desc(date))
                colnames(total_data) <- str_to_upper(colnames(total_data))
                
                # 空数据二次保护（极端兜底）
                if (nrow(total_data) == 0) stop("Merged dataset has zero rows")
                
                # ===================== 核心修复：优化CN/EN双sheet拆分逻辑 =====================
                # 1. 智能识别标题列：同时匹配英文TITLE类、中文标题类列名，优先选有内容的列
                title_col_candidates <- grep("TITLE|标题", colnames(total_data), ignore.case = TRUE, value = TRUE)
                
                if (length(title_col_candidates) == 0) {
                    title_col <- colnames(total_data)[4]
                    showNotification(paste0("⚠️ No TITLE/标题 column found, use column: ", title_col), type = "warning")
                } else {
                    non_empty_count <- sapply(title_col_candidates, function(col) sum(!is.na(total_data[[col]]) & trimws(total_data[[col]]) != ""))
                    title_col <- names(which.max(non_empty_count))
                }
                
                debug_msg <- paste0("🔍 Identified Title Column: ", title_col, " | Total Rows: ", nrow(total_data))
                message(debug_msg)
                
                # 修复Unicode转义，分步判断规避管道解析错误
                # 中文正则双反斜杠转义
                cn_regex <- "[\\u4e00-\\u9fa5]"
                title_vec <- total_data[[title_col]]
                has_cn <- str_detect(title_vec, cn_regex)
                # 原生R替换NA，抛弃replace_na
                has_cn_filled <- ifelse(is.na(has_cn), FALSE, has_cn)
                
                # 子集拆分
                cn_data <- total_data[has_cn_filled, ]
                en_data <- total_data[!has_cn_filled, ]
                
                # 空数据兜底
                if (nrow(cn_data) == 0 && nrow(en_data) == 0) {
                    en_data <- total_data
                    showNotification("⚠️ Split result empty, all data moved to EN sheet", type = "warning")
                }                
                # 4. 新建工作簿，创建CN/EN双sheet
                wb <- createWorkbook()
                addWorksheet(wb, "CN")
                addWorksheet(wb, "EN")
                
                # 为两个sheet分别应用样式+写入数据
                apply_sheet_style(wb, "CN", cn_data)
                apply_sheet_style(wb, "EN", en_data)
                
                # 写入文件、权限兜底
                output_folder <- ifelse(check_writable(default_output_folder), default_output_folder, path_expand("~/Desktop"))
                output_path <- path(output_folder, output_file)
                
                saveWorkbook(wb, output_path, overwrite = TRUE)
                # 修复：删除非法closeWorkbook，仅清理对象释放内存
                rm(wb)
                gc()
                Sys.sleep(0.5)
                
                # 自动打开结果文件
#                shell.exec(normalizePath(output_path, winslash = "/"))
                
                # 执行完成状态输出（新增双sheet行数统计+标题列调试信息）
                fail_detail <- if(length(failed_files) > 0) paste0("\nFailed Details:\n", paste(failed_files, collapse = "\n")) else ""
                final_msg <- paste0(
                    "✅ Recycle Complete\n",
                    "Save Path: ", output_path, "\n",
                    "Success File Count: ", length(success_files), "\n",
                    "Failed File Count: ", length(failed_files), "\n",
                    "Total Merged Rows: ", nrow(total_data), "\n",
                    "🔍 Identified Title Column: ", title_col, "\n",
                    "CN Sheet (Chinese Title) Rows: ", nrow(cn_data), "\n",
                    "EN Sheet (English Title) Rows: ", nrow(en_data),
                    fail_detail
                )
                output$recycle_status <- renderPrint({cat(final_msg)})
                showNotification("TLaFT Recycled Successfully", type = "message")
                
            }, error = function(e) {
                err_msg <- paste0("Error: Recycle Failed: ", e$message)
                output$recycle_status <- renderPrint({cat(err_msg)})
                showNotification(err_msg, type = "error")
            })
        })
    })
####End Recycle TLaFT####    
    # ========== Server function ============== ==========
    #  TOC render
    render_toc_table <- function() {
        req(rv$toc_filtered)
        toc_display <- rv$toc_filtered
        # 方案优化：不在主数据新增style列，改用rowCallback行回调染色，不污染原始列结构
        highlight_vec <- rv$toc_matched
        
        dt <- datatable(toc_display, 
                        rownames = FALSE, 
                        selection = "multiple",
                        escape = FALSE,
                        options = list(
                            destroy = TRUE,  # 强制销毁旧表格实例，核心！解决缓存列错位
                            pageLength = 50,
                            scrollX = TRUE,
                            scrollY = "400px",
                            dom = "lftip",
                            autoWidth = TRUE,
                            columnDefs = list(
                                list(targets = "_all", className = "dt-left")
                            ),
                            deferRender = FALSE,
                            scroller = FALSE,
                            # 行回调：直接根据行索引高亮，完全不需要新增style列
                            rowCallback = JS(sprintf("
                      function(row, data, index) {
                        const highlightArr = %s;
                        if(highlightArr[index]) {
                          $(row).css('background-color', '#d4edda');
                        }
                      }
                    ", jsonlite::toJSON(highlight_vec))
                            )
                        ))
        # 移除原formatStyle、新增/删除style列所有代码，不再依赖临时样式列
        return(dt)
    }
    
    # Render button for batch matching 
    output$toc_select_button_ui <- renderUI({
        req(rv$toc_filtered)  # 确保TOC数据加载后才显示按钮
        
        # 根据全选状态切换图标（可选优化）
        icon_name <- if(rv$is_toc_all_selected) "eraser" else "check"
        
        div(class = "toc-select-btn",
            actionButton("toggle_select_toc", NULL, 
                         class = "btn btn-xs btn-light",
                         icon = icon(icon_name),  # 动态切换图标
                         title = if(rv$is_toc_all_selected) "Cancle highlight" else "Batch Match"  # 鼠标悬浮提示
            )
        )
    })     
    
    # Export TLaFT 
    generate_export_filename <- function() {
        # 1. 获取当前工作路径（默认保存路径）
        current_wd <- getwd()
        
        # 2. 获取TOC文件名（处理空值，避免报错）
        if(is.null(input$upload_toc) || length(input$upload_toc$name) == 0) {
            toc_filename <- "unknown_toc"
        } else {
            # 去掉TOC文件的扩展名（如.xlsx/.xls）
            toc_filename <- tools::file_path_sans_ext(input$upload_toc$name)
        }
        
        # 3. 获取选定的TOC Sheet名称（处理空值）
        if(is.null(input$toc_selected_sheet) || nchar(input$toc_selected_sheet) == 0) {
            toc_sheetname <- "unknown_sheet"
        } else {
            toc_sheetname <- input$toc_selected_sheet
            # 替换Sheet名称中的特殊字符（避免文件名非法）
            toc_sheetname <- str_replace_all(toc_sheetname, "[\\\\/:*?\"<>|]", "_")
        }
        
        # 4. 拼接最终文件名：TOCexcel名称-选定的Sheet名称-title-footnote.xlsx
        final_filename <- paste0(toc_filename, "-", toc_sheetname, "-TLaFT.xlsx")
        # 5. 拼接完整路径（当前工作路径 + 生成的文件名）
        full_export_path <- file.path(current_wd, final_filename)
        
        return(list(
            wd = current_wd,          # 当前工作路径
            filename = final_filename,# 生成的文件名
            full_path = full_export_path # 完整导出路径
        ))
    } 
    
    
    
    # ========== Server Interaction ==============
    #key steps
    rv <- reactiveValues(
        x_raw = NULL,
        x_colnames = NULL,
        x_sheets = NULL,
        toc_raw = NULL,
        toc_filtered = NULL,
        toc_sheets = NULL,
        match_data = NULL,
        match_data_filtered = NULL,
        cart_data = NULL,
        toc_matched = logical(0),
        is_toc_all_selected = FALSE,
        batch_full_match_result = NULL
    )
    
    
    
    # -------------------------- X文件加载 --------------------------
    observeEvent(input$upload_x, {
        req(input$upload_x)
        runjs("$('#x_file_container').removeClass('loaded');")
        rv$x_sheets <- excel_sheets(input$upload_x$datapath)
        output$x_sheet_select_ui <- renderUI({
            selectInput("x_selected_sheet", "Select Sheet", rv$x_sheets, rv$x_sheets[1]) %>% 
                tagAppendAttributes(class = "sheet-select")
        })
    })
    
    observeEvent(input$x_selected_sheet, {
        req(input$upload_x, input$x_selected_sheet)
        rv$x_raw <- tryCatch({
            df <- read_excel(input$upload_x$datapath, sheet = input$x_selected_sheet, na = "", trim_ws = TRUE)
            df <- as.data.frame(df, stringsAsFactors = FALSE)
            colnames(df) <- toupper(make.names(colnames(df), unique = TRUE))
            for(col in colnames(df)) {
                df[[col]] <- sapply(df[[col]], flatten_col)
                df[[col]][is.na(df[[col]])] <- ""
            }
            runjs("$('#x_file_container').addClass('loaded');")
            df
        }, error = function(e) {
            runjs("$('#x_file_container').removeClass('loaded');")
            NULL
        })
        
        if(is.null(rv$x_raw)) {
            showNotification("❌ Failed to read X sheet!", type = "error")
            return()
        }
        
        rv$x_colnames <- colnames(rv$x_raw)
    })
    
    # -------------------------- TOC表格初始化 --------------------------
    output$toc_table <- renderDT({
        datatable(
            data.frame(stringsAsFactors = FALSE),  
            rownames = FALSE, 
            selection = "multiple",
            escape = TRUE,
            options = list(
                dom = "t",
                scrollX = TRUE,
                scrollY = "400px",
                autoWidth = TRUE,
                columnDefs = list(list(targets = "_all", className = "dt-left"))
            )
        )
    }, server = FALSE)
    
    # -------------------------- TOC文件加载 --------------------------
    observeEvent(input$upload_toc, {
        req(input$upload_toc)
        runjs("$('#toc_file_container').removeClass('loaded');")
        rv$toc_sheets <- excel_sheets(input$upload_toc$datapath)
        output$toc_sheet_select_ui <- renderUI({
            selectInput("toc_selected_sheet", "Select TOC Sheet", rv$toc_sheets, rv$toc_sheets[1]) %>% 
                tagAppendAttributes(class = "sheet-select")
        })
    })
    
    observeEvent(input$toc_selected_sheet, {
        req(input$upload_toc, input$toc_selected_sheet)
        rv$toc_raw <- tryCatch({
            df <- read_excel(input$upload_toc$datapath, sheet = input$toc_selected_sheet, na = "", trim_ws = TRUE)
            df <- as.data.frame(df, stringsAsFactors = FALSE)
            colnames(df) <- toupper(make.names(colnames(df), unique = TRUE))
            df$NO <- 1:nrow(df)
            for(col in colnames(df)) {
                df[[col]] <- sapply(df[[col]], flatten_col)
                df[[col]][is.na(df[[col]])] <- ""
            }
            runjs("$('#toc_file_container').addClass('loaded');")
            df
        }, error = function(e) {
            runjs("$('#toc_file_container').removeClass('loaded');")
            NULL
        })
        
        if(is.null(rv$toc_raw)) {
            showNotification("❌ Failed to read TOC sheet!", type = "error")
            return()
        }
        
        toc_core_cols <- c("NO", "SHELL_ID", "TITLE", "ANALYSIS_SET", "OUTPUT_NUMBER", "NOTE")
        intersect_cols <- intersect(toc_core_cols, colnames(rv$toc_raw))
        if(length(intersect_cols) == 0) intersect_cols <- colnames(rv$toc_raw)
        
        rv$toc_filtered <- rv$toc_raw[, intersect_cols, drop = FALSE]
        rv$toc_matched <- logical(nrow(rv$toc_filtered))
        
        output$toc_table <- renderDT({
            render_toc_table()
        }, server = FALSE)
        
        runjs("setTimeout(syncTableScroll, 100);")
    })
    
    # -------------------------- Step1: Click Batch matching for all TOC records --------------------------
    observeEvent(input$toggle_select_toc, {
        req(rv$toc_filtered)
        proxy <- dataTableProxy("toc_table")
        
        isolate({
            # 切换全选状态
            rv$is_toc_all_selected <- !rv$is_toc_all_selected
            
            if(rv$is_toc_all_selected) {
                # 全选状态：选中所有行 + 标记为已匹配（高亮）
                selectRows(proxy, 1:nrow(rv$toc_filtered))
                rv$toc_matched <- rep(TRUE, nrow(rv$toc_filtered))  # 标记所有行为已匹配（高亮）
                invalidateLater(100, session)
                batch_match_logic()  # 执行批量匹配
                showNotification("Batch Matching...", type = "default")
            } else {
                # 取消全选状态：清空选中 + 取消所有高亮
                selectRows(proxy, integer(0))
                rv$toc_matched <- rep(FALSE, nrow(rv$toc_filtered))  # 取消所有行的高亮标记
                rv$batch_full_match_result <- NULL
                rv$cart_data <- NULL
                
                showNotification("Canceled Batch Match!", type = "default")
            }
        })
        
        # 重新渲染表格，更新高亮样式
        output$toc_table <- renderDT({
            render_toc_table()
        }, server = FALSE)
        
        runjs("setTimeout(syncTableScroll, 100);")
    })
    
    # -------------------------- Step2: Execute Batch matching for all TOC records --------------------------
    batch_match_logic <- function() {
        if(is.null(rv$x_raw)) {
            showNotification("Please Upload required files first！", type = "error")
            return()
        }
        req(rv$toc_filtered)
        
        showNotification(paste0("⚙️ Matching Start, TOC Count = ", nrow(rv$toc_filtered)), type = "default")
        
        batch_cart <- list()
        matched_count <- 0
        
        for(i in 1:nrow(rv$toc_filtered)) {
            toc_data <- rv$toc_filtered[i, , drop = FALSE]
            # 安全获取标题和Shell_ID（避免空值）
            toc_title_col <- grep("^TITLE$", colnames(toc_data), ignore.case = TRUE, value = TRUE)
            toc_title <- if(length(toc_title_col) > 0) toc_data[[toc_title_col]] else ""
            toc_shell_col <- grep("^SHELL_ID$", colnames(toc_data), ignore.case = TRUE, value = TRUE)
            toc_shell_id <- if(length(toc_shell_col) > 0) toc_data[[toc_shell_col]] else ""
            
            # 安全获取X数据的标题和Shell_ID（避免空值）
            x_title_col <- grep("^TITLE1$", colnames(rv$x_raw), ignore.case = TRUE, value = TRUE)
            x_titles <- if(length(x_title_col) > 0) rv$x_raw[[x_title_col]] else character(nrow(rv$x_raw))
            x_shell_col <- grep("^SHELL_ID$", colnames(rv$x_raw), ignore.case = TRUE, value = TRUE)
            x_shell_ids <- if(length(x_shell_col) > 0) rv$x_raw[[x_shell_col]] else character(nrow(rv$x_raw))
            
            # 执行匹配
            match_res <- tryCatch({
                cosine_match_optimized(toc_title, toc_shell_id, x_titles, x_shell_ids, rv$x_raw, top_n = 1)
            }, error = function(e) {
                showNotification(paste0("Please check TOC row", e$message), type = "warning")
                return(data.frame())
            })
            
            # 核心修复：处理无匹配结果的情况
            if(nrow(match_res) > 0) {
                cart_row <- tryCatch({
                    process_cart_data(match_res[1, , drop = FALSE], toc_data)
                }, error = function(e) {
                    showNotification(paste0("Please check TOC row ", i, e$message), type = "warning")
                    # 返回空的购物车行
                    process_cart_data(data.frame(), toc_data)
                })
                batch_cart[[i]] <- cart_row
                matched_count <- matched_count + 1
                
                if(is.null(rv$batch_full_match_result)) {
                    rv$batch_full_match_result <- match_res
                } else {
                    rv$batch_full_match_result <- rbind(rv$batch_full_match_result, match_res)
                }
                
            } else {
                # 无匹配结果时创建空行（确保结构一致）
                cart_row <- process_cart_data(data.frame(), toc_data)
                batch_cart[[i]] <- cart_row
            }
            runjs("showSidebar();")  # 确保右侧面板显示
        }
        
        # 合并批量购物车数据（安全处理空列表）
        if(length(batch_cart) > 0 && !all(sapply(batch_cart, is.null))) {
            # 过滤掉空值
            batch_cart_clean <- Filter(function(x) !is.null(x) && is.data.frame(x) && nrow(x) > 0, batch_cart)
            
            if(length(batch_cart_clean) > 0) {
                # 统一列名
                all_cols <- unique(unlist(lapply(batch_cart_clean, colnames)))
                batch_cart_clean <- lapply(batch_cart_clean, function(df) {
                    for(col in all_cols) {
                        if(!col %in% colnames(df)) {
                            df[[col]] <- ""
                        }
                    }
                    df[, all_cols, drop = FALSE]
                })
                rv$cart_data <- do.call(rbind, batch_cart_clean)
            } else {
                rv$cart_data <- data.frame(TOC_No = character(0), stringsAsFactors = FALSE)
            }
        } else {
            rv$cart_data <- data.frame(TOC_No = character(0), stringsAsFactors = FALSE)
        }
        
        # 更新匹配状态
        rv$toc_matched <- rep(TRUE, nrow(rv$toc_filtered))
        showNotification(paste0("Batch Matching Completed!"), type = "default")
        
        if(!is.null(rv$batch_full_match_result) && nrow(rv$batch_full_match_result) > 0) {
            output$match_result_table <- renderRHandsontable({
                ht <- rhandsontable(rv$batch_full_match_result, width = "100%", rowHeaders = NULL) %>%
                    hot_cols(readOnly = TRUE, align = "left", fontSize = 11, stretchH = "all") %>%
                    hot_table(highlightRow = TRUE, selectionMode = "row", allowDeselect = TRUE, outsideClickDeselect = FALSE)
                # 保留原有右键Add to Cart菜单（兼容单选交互）
                ht$x$contextMenu <- list(
                    items = list(
                        addCart = list(
                            name = "Add to Cart",
                            callback = JS("function(key, selection) {
                          const hot = this;
                          let targetRowIndex;
                        
                          // 情况1：已有选中行，直接使用
                          if (selection && selection.length > 0 && Array.isArray(selection[0]) && selection[0][0] !== undefined) {
                            targetRowIndex = selection[0][0];
                          } else {
                            // 情况2：无选中，直接读取右键点击的原生行（菜单自带点击位置，稳定无卡顿）
                            const menuClickRow = hot.getSelectedLast() ? hot.getSelectedLast()[0] : null;
                            if (menuClickRow !== null && menuClickRow >= 0) {
                              targetRowIndex = menuClickRow;
                              hot.selectRows(targetRowIndex, targetRowIndex);
                            } else {
                              alert('请将鼠标放置在匹配数据行上再右键');
                              return false;
                            }
                          }
                          // 安全推送，无延时阻塞
                          Shiny.setInputValue('rightclick_add_cart_row', targetRowIndex, {priority: 'event'});
                        }")
                        )
                    )
                )
                ht
            })
        }
        
        # 重新渲染TOC表格
        output$toc_table <- renderDT({
            render_toc_table()
        }, server = FALSE)
        
        runjs("setTimeout(syncTableScroll, 100);")
    }
    
    
    # -------------------------- Step3: Match with TLaFT --------------------------
    observeEvent(input$toc_table_rows_selected, {
        req(rv$x_raw, rv$toc_filtered)
        selected_row <- input$toc_table_rows_selected
        
        if(length(selected_row) > 1) {
            selected_row <- selected_row[1]
            #showNotification("⚠️ Only single row is supported", type = "warning")
        }
        
        if(is.null(selected_row) || selected_row < 1 || selected_row > nrow(rv$toc_filtered)) {
            showNotification("❌ Invalid TOC row selected!", type = "warning")
            return()
        }
        
        toc_data <- rv$toc_filtered[selected_row, , drop = FALSE]
        # 安全获取标题和Shell_ID
        toc_title_col <- grep("^TITLE$", colnames(toc_data), ignore.case = TRUE, value = TRUE)
        toc_title <- if(length(toc_title_col) > 0) toc_data[[toc_title_col]] else ""
        toc_shell_col <- grep("^SHELL_ID$", colnames(toc_data), ignore.case = TRUE, value = TRUE)
        toc_shell_id <- if(length(toc_shell_col) > 0) toc_data[[toc_shell_col]] else ""
        
        # 安全获取X数据的标题和Shell_ID
        x_title_col <- grep("^TITLE1$", colnames(rv$x_raw), ignore.case = TRUE, value = TRUE)
        x_titles <- if(length(x_title_col) > 0) rv$x_raw[[x_title_col]] else character(nrow(rv$x_raw))
        x_shell_col <- grep("^SHELL_ID$", colnames(rv$x_raw), ignore.case = TRUE, value = TRUE)
        x_shell_ids <- if(length(x_shell_col) > 0) rv$x_raw[[x_shell_col]] else character(nrow(rv$x_raw))
        
        rv$match_data <- tryCatch({
            cosine_match_optimized(toc_title, toc_shell_id, x_titles, x_shell_ids, rv$x_raw)
        }, error = function(e) {
            showNotification(paste0("⚠️ Error in matching: ", e$message), type = "warning")
            return(data.frame())
        })
        
        # 匹配结果表格渲染完整修正代码
        output$match_result_table <- renderRHandsontable({
            if(nrow(rv$match_data) == 0) {
                df_empty <- data.frame(
                    Message = "⚠️ No matching records found (score ≥ 0.7)",
                    stringsAsFactors = FALSE
                )
                ht <- rhandsontable(df_empty, width = "100%", rowHeaders = NULL) %>% 
                    hot_cols(
                        readOnly = TRUE,
                        align = "center",
                        fontSize = 12,
                        color = "#dc3545"
                    ) %>%
                    hot_table(
                        highlightCol = FALSE, 
                        highlightRow = FALSE,
                        rowHeaders = FALSE,
                        selectable = TRUE
                    )
                # 清空空白表格右键菜单
                ht$x$contextMenu <- list()
                return(ht)
            } else {
                # 1. 先基础渲染表格
                ht <- rhandsontable(rv$match_data, width = "100%", rowHeaders = NULL) %>% 
                    hot_cols(
                        readOnly = TRUE,
                        align = "left",
                        fontSize = 11,
                        stretchH = "all"
                    ) %>%
                    hot_table(
                        highlightCol = FALSE,
                        highlightRow = TRUE,
                        rowHeaders = FALSE,
                        selectable = TRUE,
                        autoWrapRow = TRUE,
                        height = NULL,
                        selectionMode = "row",
                        allowDeselect = TRUE,
                        outsideClickDeselect = FALSE
                    ) %>%
                    # 仅挂载自定义菜单项（不依赖覆盖默认）
                    hot_context_menu(
                        customOpts = list(
                            addCart = list(
                                name = "Add to Cart",
                                callback = JS("function(key, selection) {
              if (!selection || selection.length === 0) {
                alert('Please select the aim row');
                return;
              }
              const finalRow = selection[0][0];
              Shiny.setInputValue('rightclick_add_cart_row', finalRow, {priority: 'event'});
            }")
                            )
                        ),
                        allowRowEdit = FALSE,  # 禁用行编辑（插入/删除行彻底锁死）
                        allowColEdit = FALSE  # 禁用列编辑
                    )
                ht$x$contextMenu <- list(
                    items = list(
                        addCart = list(
                            name = "Add to Cart",
                            callback = JS("function(key, selection) {
                          const hot = this;
                          let targetRowIndex;

                          // 情况1：已有选中行，直接使用
                          if (selection && selection.length > 0 && Array.isArray(selection[0]) && selection[0][0] !== undefined) {
                            targetRowIndex = selection[0][0];
                          } else {
                            // 情况2：无选中，直接读取右键点击的原生行（菜单自带点击位置，稳定无卡顿）
                            const menuClickRow = hot.getSelectedLast() ? hot.getSelectedLast()[0] : null;
                            if (menuClickRow !== null && menuClickRow >= 0) {
                              targetRowIndex = menuClickRow;
                              hot.selectRows(targetRowIndex, targetRowIndex);
                            } else {
                              alert('请将鼠标放置在匹配数据行上再右键');
                              return false;
                            }
                          }
  // 安全推送，无延时阻塞
                          Shiny.setInputValue('rightclick_add_cart_row', targetRowIndex, {priority: 'event'});
}")
                        )
                    )
                )
                return(ht)
            }
        })
        
    })
    
    # -------------------------- Step4: Match into CART --------------------------
    # 替换原来的 observeEvent(input$match_result_table_rows_selected, {...})
    observeEvent(input$rightclick_add_cart_row, {
        req(
            rv$match_data, 
            rv$toc_filtered, 
            input$toc_table_rows_selected, 
            input$rightclick_add_cart_row
        )
        
        # 从右键JS获取0起始行号，+1转为R数据框下标
        selected_match_row <- input$rightclick_add_cart_row + 1
        selected_toc_row <- input$toc_table_rows_selected
        if(length(selected_toc_row) > 1) selected_toc_row <- selected_toc_row[1]
        
        x_match_row <- rv$match_data[selected_match_row, , drop = FALSE]
        toc_data <- rv$toc_filtered[selected_toc_row, , drop = FALSE]
        
        # 后续购物车组装逻辑完全原样保留
        cart_row <- tryCatch({
            process_cart_data(x_match_row, toc_data)
        }, error = function(e) {
            showNotification(paste0("⚠️ Error in addting to cart: ", e$message), type = "warning")
            process_cart_data(data.frame(), toc_data)
        })
        
        if(is.null(rv$cart_data) || nrow(rv$cart_data) == 0) {
            rv$cart_data <- cart_row
        } else {
            toc_no_val <- cart_row$TOC_No[1]
            if(toc_no_val %in% rv$cart_data$TOC_No) {
                common_cols <- intersect(colnames(rv$cart_data), colnames(cart_row))
                rv$cart_data[rv$cart_data$TOC_No == toc_no_val, common_cols] <- cart_row[, common_cols]
            } else {
                new_cols <- setdiff(colnames(cart_row), colnames(rv$cart_data))
                for(col in new_cols) rv$cart_data[[col]] <- ""
                old_cols <- setdiff(colnames(rv$cart_data), colnames(cart_row))
                for(col in old_cols) cart_row[[col]] <- ""
                cart_row <- cart_row[, colnames(rv$cart_data), drop = FALSE]
                rv$cart_data <- rbind(rv$cart_data, cart_row)
            }
        }
        
        rv$toc_matched[selected_toc_row] <- TRUE
        showNotification(paste0("✅ TOC No.", selected_toc_row, " added to cart!"), type = "default")
        
        output$toc_table <- renderDT({
            render_toc_table()
        }, server = FALSE)
        
        runjs("setTimeout(syncTableScroll, 100);")
    })
    
    #--------------------------- Step5: Render CART----------- 
    output$cart_table <- renderRHandsontable({
        req(rv$cart_data)
        
        ## ---------- 1. 空购物车提示 ----------
        if (!is.data.frame(rv$cart_data) || nrow(rv$cart_data) == 0) {
            df_empty <- data.frame(
                Message = "🛒 Cart is empty (add TOC matches to populate)",
                stringsAsFactors = FALSE
            )
            return(
                rhandsontable(df_empty, width = "100%", rowHeaders = NULL) %>%
                    hot_cols(readOnly = TRUE,
                             align  = "center",
                             fontSize = 11,
                             color  = "#6c757d") %>%
                    hot_table(highlightCol = FALSE, highlightRow = FALSE, rowHeaders = FALSE)
            )
        }
        
        ## ---------- 2. 数据预处理 ----------
        rv$cart_data[is.na(rv$cart_data)] <- ""
        
        ## 需要高亮/只读的列名
        highlight_col_names <- c(names(COLUMN_MAPPING), "TOC_No")
        
        ## ---------- 3. 构建基础表 ----------
        hot <- rhandsontable(rv$cart_data, width = "100%", rowHeaders = NULL) %>%
            hot_table(highlightCol = FALSE,
                      highlightRow = FALSE, 
                      rowHeaders = FALSE,
                      stretchH = "all")
        
        ## ---------- 4. 逐列设置：只读 + 渲染器 ----------
        for (j in seq_len(ncol(rv$cart_data))) {
            col_name <- colnames(rv$cart_data)[j]
            is_locked <- col_name %in% highlight_col_names   # TRUE=只读
            
            # 该列专用渲染器
            renderer_js <- sprintf(
                "function(instance, td, row, col, prop, value, cellProperties){
         Handsontable.renderers.TextRenderer.apply(this, arguments);
         if(%s){                       // 高亮列
           td.style.backgroundColor = '#e7f5ff';
           td.style.fontWeight = '500';
         }else{
           td.style.backgroundColor = '';
           td.style.fontWeight = 'normal';
         }
       }", tolower(is_locked))
            
            hot <- hot %>% hot_col(
                col      = j,
                readOnly = is_locked,   # ✅ 逐列生效，非高亮列=FALSE 可编辑
                renderer = renderer_js,
                align    = "center",
                fontSize = 11,
                stretchH = "all"
            )
        }
        
        ## ---------- 5. 返回最终表 ----------
        hot
    })  
    # -------------------------- Step6: Edit CART --------------------------
    observeEvent(input$cart_table, {
        req(rv$cart_data)
        # 读取rhandsontable的编辑后数据
        hot_data <- hot_to_r(input$cart_table)
        
        # 只在数据变化时更新（避免重复触发）
        if(!identical(hot_data, rv$cart_data)) {
            rv$cart_data <- hot_data
            showNotification("✅ Data Refreshed!", type = "default")
        }
    })
    
    # -------------------------- Step7: Clear CART --------------------------
    observeEvent(input$clear_cart, {
        # 安全清空，避免空值报错
        rv$cart_data <- data.frame(TOC_No = character(0), stringsAsFactors = FALSE)
        showNotification("✅ Cart cleared! ", type = "default")
    })
    # -------------------------- Step8: Export TLaFT --------------------------
    observeEvent(input$export_z, {
        # 防护1：校验购物车数据
        if(is.null(rv$cart_data) || nrow(rv$cart_data) == 0) {
            showNotification("⚠️ Warning: No record in Cart！", type = "warning")
            return()
        }
        
        # 防护2：调用自动生成文件名函数
        export_info <- generate_export_filename()
        full_export_path <- export_info$full_path
        
        tryCatch({
            # 准备导出数据
            export_data <- rv$cart_data
            export_data[is.na(export_data)] <- ""
            
            # 写入Excel
            wb <- createWorkbook()
            addWorksheet(wb, "Title_Footnote")
            writeData(wb, "Title_Footnote", export_data, rowNames = FALSE)
            apply_sheet_style(wb, "Title_Footnote", export_data)
            saveWorkbook(wb, full_export_path, overwrite = TRUE)
            shell.exec(normalizePath(full_export_path, winslash = "/"))            
            # 导出成功反馈（显示完整路径）
            output$export_status <- renderPrint({
                cat(paste0("Export Successed!\n",
                           "Pathway: ",  full_export_path, "\n",
                           "Records count: ", nrow(export_data), "\n",
                           "Total columns: ", ncol(export_data)))
            })
            showNotification(paste0("✅ Export Successed!"),
                             type = "message")
            
        }, error = function(e) {
            # 异常处理：详细报错
            showNotification(paste0("⚠️ Export Failed: ", e$message), type = "error")
            output$export_status <- renderPrint({
                cat(paste0("Export Failed:\n",
                           "Error: ", e$message, "\n"))
            })
        })
    })
}

# ===================== 6. app calling =====================
shinyApp(ui = ui, server = server)
