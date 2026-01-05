#### Title-Footnote Excel 筛选导出工具（整合所有修复+优化，可直接运行）
#### 依赖包安装（Windows本地，无需管理员权限，国内镜像快速安装）
options(repos = c(CRAN = "https://mirrors.tuna.tsinghua.edu.cn/CRAN/"))
# 安装所有必需依赖包
install.packages(c("shiny", "shinyFiles", "readxl", "openxlsx", "fs", "dplyr", "stringr", "DT"))

#### 加载核心库
library(shiny)
library(shinyFiles)       # 文件夹选择控件
library(readxl)           # Excel文件读取（兼容xlsx/xls/xlsm）
library(openxlsx)         # Excel文件导出
library(fs)               # 安全文件/目录操作
library(dplyr)            # 数据处理
library(stringr)          # 字符串筛选
library(DT)               # 带勾选框的数据表格

#### ===================== 核心工具函数（整合所有优化：解除限制+兼容+日志）=====================
# 1. 解析Shiny文件夹选择路径（兼容Windows反斜杠+特殊字符+空格）
parse_shiny_dir <- function(dir_input) {
  if (is.null(dir_input) || !is.list(dir_input) || is.null(dir_input$path)) {
    return(NULL)
  }
  path_parts <- unlist(dir_input$path, use.names = FALSE)
  path_parts <- path_parts[path_parts != "" & !is.na(path_parts)]
  if (length(path_parts) == 0) {
    return(NULL)
  }
  # 转换为Windows规范路径，增强特殊字符兼容
  full_path <- paste(path_parts, collapse = .Platform$file.sep)
  full_path <- fs::path_norm(full_path)
  # 处理路径首尾空格
  full_path <- str_trim(full_path)
  return(full_path)
}

# 2. 路径分隔符函数（兼容跨系统）
path_sep <- function() {
  return(.Platform$file.sep)
}

# 3. 统一数据列类型（避免数据格式冲突）
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

# 4. 检查目录可写性（Windows系统目录规避）
check_writable <- function(dir) {
  if (is.null(dir) || dir == "" || !dir_exists(dir)) {
    return(FALSE)
  }
  # 跳过Windows系统保护目录
  if (str_detect(dir, "^[CDE]:/Windows/")) {
    return(FALSE)
  }
  # 临时文件测试可写性，自动清理
  test_file <- tempfile(tmpdir = dir)
  writable <- tryCatch({
    file.create(test_file); file.remove(test_file); TRUE
  }, error = function(e) FALSE)
  return(writable)
}

# 5. 检查目录可访问性（放宽限制，仅跳过系统目录）
check_dir_accessible <- function(dir_path) {
  if (is.null(dir_path) || dir_path == "" || !dir_exists(dir_path)) {
    return(FALSE)
  }
  # 仅跳过Windows系统保护目录，放宽用户文件夹限制
  if (str_detect(dir_path, "^[CDE]:/Windows/|^[CDE]:/Program Files/|^[CDE]:/Program Files (x86)/")) {
    return(FALSE)
  }
  tryCatch({
    list.files(dir_path, recursive = FALSE)
    TRUE
  }, error = function(e) {
    warning("⚠️ 无访问权限，跳过文件夹: ", dir_path)
    FALSE
  })
}

# 6. 安全列出Excel文件（核心优化：10层深度+隐藏文件+xlsm+扫描日志）
safe_list_target_excels <- function(root_dir, keyword, max_depth = 10) {
  if (is.null(root_dir) || root_dir == "") {
    return(character(0))
  }
  
  valid_excel_files <- character(0)
  # 递归获取所有子文件夹（默认10层，可按需调整）
  all_folders <- list.dirs(path = root_dir, recursive = TRUE, full.names = TRUE)
  # 过滤超出最大深度的文件夹（删除以下3行可完全取消深度限制）
  folder_depths <- sapply(all_folders, function(x) {
    length(str_split(x, path_sep())[[1]]) - length(str_split(root_dir, path_sep())[[1]])
  })
  all_folders <- all_folders[folder_depths <= max_depth]
  all_folders <- unique(c(root_dir, all_folders))
  
  # 输出扫描日志（控制台查看，方便排查）
  message("🔍 正在扫描文件夹（共", length(all_folders), "个，最大深度", max_depth, "层）：")
  for (dir in all_folders) {
    message("   - ", dir)
  }
  
  for (dir in all_folders) {
    if (!check_dir_accessible(dir)) next
    
    dir_excels <- tryCatch({
      # 核心优化：all=TRUE（读取隐藏文件）、新增.xlsm（宏文件）支持
      all_excels <- dir_ls(
        path = dir,
        glob = c("*.xlsx", "*.xls", "*.xlsm"),
        recurse = FALSE,
        ignore.case = TRUE,
        all = TRUE
      )
      # 关键词筛选：仅当关键词非空时过滤（留空不筛选）
      if (!is.null(keyword) && keyword != "" && str_trim(keyword) != "") {
        all_excels <- all_excels[str_detect(basename(all_excels), regex(str_trim(keyword), ignore_case = TRUE))]
      }
      all_excels
    }, error = function(e) {
      warning("⚠️ 无法读取文件夹: ", dir, "，错误: ", e$message)
      character(0)
    })
    
    valid_excel_files <- unique(c(valid_excel_files, dir_excels))
  }
  
  # 输出扫描结果日志（控制台查看）
  if (length(valid_excel_files) > 0) {
    message("✅ 扫描到", length(valid_excel_files), "个符合条件的Excel文件：")
    for (file in valid_excel_files) {
      message("   - ", file)
    }
  } else {
    message("❌ 未扫描到任何符合条件的Excel文件")
  }
  
  return(valid_excel_files)
}

# 7. 提取2级文件夹结构（用于数据分类）
extract_two_level_folders <- function(file_path, root_dir) {
  if (is.null(file_path) || is.null(root_dir) || root_dir == "") {
    return(list(Protocol = "", Task = ""))
  }
  rel_path <- path_rel(file_path, root_dir)
  path_parts <- str_split(rel_path, path_sep(), simplify = TRUE)[1, ]
  folder_levels <- path_parts[-length(path_parts)]
  folder_levels <- folder_levels[folder_levels != "" & !is.na(folder_levels)]
  
  protocol <- ifelse(length(folder_levels) >= 1, folder_levels[1], "")
  task <- ifelse(length(folder_levels) >= 2, folder_levels[2], "")
  
  return(list(Protocol = protocol, Task = task))
}

# 8. 读取并整合所有Excel数据（30秒超时保护+容错处理）
load_all_data <- function(target_folder, keyword, timeout = 30) {
  if (is.null(target_folder) || target_folder == "") {
    return(NULL)
  }
  
  # 超时保护：超过30秒自动终止，避免R进程阻塞
  result <- tryCatch({
    setTimeLimit(elapsed = timeout, transient = TRUE)
    # 获取符合条件的Excel文件
    excel_files <- safe_list_target_excels(root_dir = target_folder, keyword = keyword)
    if (length(excel_files) == 0) return(NULL)
    
    all_data_list <- list()
    for (file_path in excel_files) {
      file_name <- path_file(file_path)
      
      tryCatch({
        # 兼容.xlsx/.xls/.xlsm格式读取
        df <- read_excel(
          path = file_path,
          sheet = 1,          
          na = "",            
          trim_ws = TRUE      
        )
        
        if (nrow(df) == 0) {
          warning("⚠️ 文件无有效数据，跳过: ", file_name)
          next
        }
        
        # 统一列类型 + 添加分类/时间信息
        df <- unify_col_type(df)
        folder_info <- extract_two_level_folders(file_path, target_folder)
        file_mod_time <- file_info(file_path)$modification_time
        timestamp_dt <- as.Date(file_mod_time, tz = "UTC")
        
        df_with_all_cols <- df %>% 
          mutate(
            Protocol = folder_info$Protocol,
            Task = folder_info$Task,
            TLaFT_XLSX = file_name,
            Timestamp = timestamp_dt
          ) %>% 
          relocate(Protocol, .before = 1) %>% 
          relocate(Task, .after = Protocol) %>% 
          relocate(TLaFT_XLSX, .after = Task) %>% 
          relocate(Timestamp, .after = TLaFT_XLSX)
        
        all_data_list[[file_path]] <- df_with_all_cols
      }, error = function(e) {
        warning("⚠️ 读取文件失败，跳过: ", file_path, "，错误: ", e$message)
      })
    }
    
    if (length(all_data_list) == 0) return(NULL)
    # 合并所有数据并添加行号（方便用户查看）
    total_data <- bind_rows(all_data_list)
    total_data <- total_data %>% mutate(行号 = row_number(), .before = 1)
    return(total_data)
  }, error = function(e) {
    warning("⚠️ 数据加载超时（超过", timeout, "秒）或出错: ", e$message)
    return(NULL)
  }, finally = {
    # 重置时间限制，不影响后续操作
    setTimeLimit(elapsed = Inf, transient = TRUE)
  })
  
  return(result)
}

# 9. 导出选中记录为Excel（保留样式+兼容Windows路径+自动兜底）
export_selected_data <- function(selected_data, output_path) {
  if (is.null(selected_data) || nrow(selected_data) == 0 || is.null(output_path)) {
    return(FALSE)
  }
  
  # 移除行号列，保留业务数据
  export_data <- selected_data %>% select(-`行号`)
  
  # 创建工作簿并设置美观样式
  wb <- createWorkbook()
  sheet_name <- "title-footnote"
  addWorksheet(wb, sheetName = sheet_name)
  
  # 表头样式（灰色背景+粗体+边框）
  header_style <- createStyle(
    fill = createFill(fgColor = "#EDF2F9"),
    bgFill = "#EDF2F9",
    fontName = "Arial",
    fontSize = 10,
    fontColour = "#000000",
    textDecoration = "bold",
    border = "TopBottomLeftRight",
    borderColour = "#CCCCCC"
  )
  
  # 日期列样式（yyyy-mm-dd格式）
  date_style <- createStyle(numFmt = "yyyy-mm-dd")
  
  # 写入数据并添加筛选器
  writeData(
    wb,
    sheet = sheet_name,
    x = export_data,
    startRow = 1,
    withFilter = TRUE,
    headerStyle = header_style,
    colNames = TRUE
  )
  
  # 设置列宽（日期列固定15，其他自动适配）
  timestamp_col_idx <- which(colnames(export_data) == "Timestamp")
  col_widths <- rep("auto", ncol(export_data))
  if (length(timestamp_col_idx) > 0) {
    col_widths[timestamp_col_idx] <- 15
  }
  setColWidths(wb, sheet = sheet_name, cols = 1:ncol(export_data), widths = col_widths)
  
  # 应用日期样式到Timestamp列
  if (length(timestamp_col_idx) > 0) {
    addStyle(
      wb,
      sheet = sheet_name,
      style = date_style,
      rows = 2:(nrow(export_data) + 1),
      cols = timestamp_col_idx,
      gridExpand = TRUE,
      stack = TRUE
    )
  }
  
  # 保存文件（自动覆盖已有文件，兼容Windows特殊路径）
  tryCatch({
    saveWorkbook(wb, output_path, overwrite = TRUE)
    return(TRUE)
  }, error = function(e) {
    warning("⚠️ 导出失败，错误: ", e$message)
    return(FALSE)
  })
}

#### ===================== Shiny UI界面（整合优化：默认关键词留空+提示清晰）=====================
ui <- fluidPage(
  titlePanel("📊 Excel 筛选导出工具（整合所有修复，可直接运行）"),
  tags$hr(),
  
  # 第一行：参数输入区（文件夹选择+关键词+导出文件名+按钮）
  fluidRow(
    column(4,
           # 文件夹选择按钮（美化样式，宽100%）
           tagAppendAttributes(
             shinyDirButton("target_folder", "选择目标文件夹", "请选择包含Excel文件的根文件夹"),
             style = "width:100%; margin-bottom:10px; background-color:#007bff; color:white; border:none; border-radius:4px;"
           ),
           # 选中文件夹路径显示（美化样式）
           verbatimTextOutput("selected_folder") %>% 
             tagAppendAttributes(style = "font-size:12px; color:#666; border:1px solid #eee; padding:5px; border-radius:4px; min-height:60px;")
    ),
    column(3,
           # 关键词输入框（默认留空，避免筛选遗漏）
           tagAppendAttributes(
             textInput(
               inputId = "keyword",
               label = "文件筛选关键词（留空读取所有Excel，忽略大小写）",
               value = "", # 默认留空，核心优化点
               placeholder = "例：title-footnote（输入后仅匹配包含该字符串的文件）"
             ),
             style = "width:100%; padding:5px; border-radius:4px;"
           ),
           helpText("支持格式：.xlsx / .xls / .xlsm，可读取隐藏文件")
    ),
    column(3,
           # 导出文件名输入框（自动补充.xlsx后缀）
           tagAppendAttributes(
             textInput(
               inputId = "output_file",
               label = "导出文件名",
               value = "title-footnote.xlsx",
               placeholder = "例如：result.xlsx（需带.xlsx后缀）"
             ),
             style = "width:100%; padding:5px; border-radius:4px;"
           ),
           helpText("提示：无.xlsx后缀将自动补充")
    ),
    column(2,
           # 加载/导出按钮（垂直排列，美化样式）
           actionButton("load_data", "📥 加载数据", class = "btn-primary", style = "width:100%; margin-top:25px; border-radius:4px;"),
           actionButton("export_data", "📤 导出选中", class = "btn-success", style = "width:100%; margin-top:10px; border-radius:4px;")
    )
  ),
  
  tags$hr(),
  
  # 第二行：数据操作区（全选/取消全选 + 带勾选框的DT表格）
  fluidRow(
    column(12,
           h4("📋 数据列表（勾选需要导出的记录）"),
           # 全选/取消全选按钮
           div(
             actionButton("select_all", "✅ 全选所有记录", class = "btn-default", style = "margin-right:10px; border-radius:4px;"),
             actionButton("deselect_all", "❌ 取消全选", class = "btn-default", style = "margin-right:10px; border-radius:4px;"),
             style = "margin-bottom:15px;"
           ),
           # 数据表格（支持滚动、导出、中文显示）
           DTOutput("data_table"),
           style = "margin-bottom:20px;"
    )
  ),
  
  tags$hr(),
  
  # 第三行：导出结果 + 最近导出记录
  fluidRow(
    column(8,
           h4("📤 导出结果"),
           verbatimTextOutput("export_result") %>% 
             tagAppendAttributes(style = "font-size:14px; padding:10px; border-radius:5px; border:1px solid #eee; min-height:80px;")
    ),
    column(4,
           h4("📜 最近导出记录（最多5条）"),
           verbatimTextOutput("export_history") %>% 
             tagAppendAttributes(style = "font-size:12px; padding:10px; border-radius:5px; background-color:#f8f9fa; max-height:200px; overflow-y:auto; border:1px solid #eee;")
    )
  ),
  
  # 全局样式优化（Windows界面友好，美化显示）
  tags$style(HTML("
    .dataTables_wrapper .dataTables_length {margin-bottom:10px;}
    .dataTables_checkbox {width:20px; height:20px;}
    .btn {padding:6px 12px; font-size:14px;}
    .help-text {font-size:11px; color:#999; margin-top:5px;}
    hr {border-top:1px solid #ddd; margin:20px 0;}
    .btn-default {background-color:#e9ecef; border-color:#ddd;}
    body {padding:20px; font-family:Microsoft YaHei; background-color:#fafbfc;}
    .container {max-width:1200px; margin:0 auto;}
    /* 修复通知遮罩层级问题，避免被遮挡 */
    .shiny-notification {z-index: 9999 !important;}
  "))
)

#### ===================== Shiny Server逻辑（整合所有修复：兜底+容错+无报错）=====================
server <- function(input, output, session) {
  # 核心修复：移除restrictions参数，解除目录访问限制，支持xlsm格式
  shinyDirChoose(
    input = input,
    id = "target_folder",
    roots = c(
      "桌面" = fs::path_expand("~/Desktop"),
      "我的文档" = fs::path_expand("~/Documents"),
      "此电脑（C盘）" = "C:/",
      "用户目录" = fs::path_expand("~")
    ),
    filetypes = c("", "xlsx", "xls", "xlsm"), # 支持xlsm格式
    session = session
  )
  
  # 1. 显示选中的文件夹路径（带提示信息）
  output$selected_folder <- renderPrint({
    target_path <- parse_shiny_dir(input$target_folder)
    if (is.null(target_path)) {
      cat("👉 请点击上方按钮选择目标文件夹（推荐：桌面/我的文档）\n")
      cat("📌 提示：仅访问你授权的目录，无其他目录扫描")
    } else {
      cat("✅ 当前选中文件夹：\n")
      cat(target_path, "\n")
      cat("📌 提示：将扫描该目录下10层内的.xlsx/.xls/.xlsm文件（含隐藏文件）")
    }
  })
  
  # 2. 响应式数据存储（所有数据本地存储，无网络传输）
  all_data <- reactiveVal(NULL)       # 存储加载的所有Excel数据
  export_history <- reactiveVal(list()) # 存储最近5条导出记录
  
  # 3. 加载数据按钮逻辑（核心修复：非法type+兜底移除通知+超时保护）
  observeEvent(input$load_data, {
    target_path <- parse_shiny_dir(input$target_folder)
    keyword <- trimws(input$keyword) # 自动去除关键词前后空格
    
    # 输入校验：未选择文件夹则提示
    if (is.null(target_path) || target_path == "") {
      showNotification("⚠️ 请先选择有效的目标文件夹！", type = "warning", duration = 5)
      return()
    }
    
    # 显示加载通知（合法type=message，隐藏关闭按钮）
    showNotification(
      "🔄 正在扫描并加载Excel数据，请稍候...（最多30秒）", 
      type = "message",
      duration = NULL, 
      id = "load_notif",
      closeButton = FALSE
    )
    
    # tryCatch + finally 兜底：无论成功/失败/超时，都移除加载通知
    tryCatch({
      # 加载数据（带30秒超时保护）
      data <- load_all_data(target_path, keyword, timeout = 30)
      if (is.null(data)) {
        all_data(NULL)
        showNotification("❌ 未找到符合条件的Excel文件！可尝试留空关键词或更换文件夹", type = "warning", duration = 8)
      } else {
        all_data(data)
        file_count <- length(unique(data$TLaFT_XLSX))
        showNotification(paste0("✅ 成功加载 ", nrow(data), " 条记录（", file_count, " 个Excel文件）"), type = "message", duration = 5)
      }
    }, error = function(e) {
      all_data(NULL)
      showNotification(paste0("❌ 加载失败：", e$message), type = "error", duration = 10)
    }, finally = {
      # 强制移除加载通知，解决页面变暗问题
      removeNotification(id = "load_notif")
    })
  })
  
  # 4. 全选/取消全选逻辑（修复非法type，提示正常显示）
  observeEvent(input$select_all, {
    req(all_data()) # 依赖：必须先加载数据
    all_rows <- 1:nrow(all_data())
    updateDT(session, "data_table", selection = list(selected = all_rows))
    showNotification("✅ 已选中所有记录！", type = "message", duration = 3)
  })
  
  observeEvent(input$deselect_all, {
    req(all_data()) # 依赖：必须先加载数据
    updateDT(session, "data_table", selection = list(selected = integer(0)))
    showNotification("❌ 已取消所有记录的选中状态！", type = "message", duration = 3)
  })
  
  # 5. 渲染带勾选框的数据表格（优化渲染速度，支持中文+滚动）
  output$data_table <- renderDT({
    req(all_data()) # 依赖：必须先加载数据
    datatable(
      data = all_data(),
      selection = list(mode = "multiple", target = "row", selected = NULL),
      rownames = FALSE,
      extensions = c("Scroller", "Buttons"),
      options = list(
        pageLength = 10,
        lengthMenu = c(10, 20, 50, 100),
        scrollX = TRUE,
        scrollY = "500px",
        scroller = TRUE,
        dom = "lBfrtip",
        buttons = c("csv", "excel", "print"),
        language = list(
          url = "//cdn.datatables.net/plug-ins/1.13.4/i18n/zh-CN.json"
        ),
        autoWidth = FALSE # 禁用自动宽度，提升渲染速度
      ),
      class = "display compact table-striped table-hover"
    ) %>% 
      formatDate("Timestamp", method = "toLocaleDateString", params = list("zh-CN")) # 日期格式化显示
  })
  
  # 6. 导出选中数据逻辑（自动兜底到桌面+历史记录+样式保留）
  observeEvent(input$export_data, {
    req(all_data(), input$output_file) # 依赖：加载数据+输入导出文件名
    selected_rows <- input$data_table_rows_selected
    target_path <- parse_shiny_dir(input$target_folder)
    
    # 校验1：未选中记录
    if (length(selected_rows) == 0) {
      showNotification("⚠️ 请先勾选需要导出的记录！", type = "warning", duration = 5)
      output$export_result <- renderPrint({
        cat("❌ 导出失败：未选中任何记录！")
      })
      return()
    }
    
    # 校验2：未选择文件夹
    if (is.null(target_path) || target_path == "") {
      showNotification("⚠️ 请先选择有效的目标文件夹！", type = "error", duration = 5)
      output$export_result <- renderPrint({
        cat("❌ 导出失败：未选择目标文件夹！")
      })
      return()
    }
    
    # 处理导出文件名：自动补充.xlsx后缀
    output_file <- trimws(input$output_file)
    if (!str_ends(output_file, ".xlsx")) {
      output_file <- paste0(output_file, ".xlsx")
    }
    
    # 构建导出路径：默认保存到 选中文件夹/docs/studymgt/
    output_path <- path(target_path, "docs", "studymgt", output_file)
    output_dir <- path_dir(output_path)
    
    # 创建导出目录（不存在则自动递归创建）
    if (!dir_exists(output_dir)) {
      dir_create(output_dir, recursive = TRUE)
    }
    
    # 权限兜底：目标目录不可写时，自动导出到桌面
    if (!check_writable(output_dir)) {
      output_path <- path(fs::path_expand("~/Desktop"), output_file)
      showNotification("⚠️ 目标目录不可写，将自动导出到桌面！", type = "warning", duration = 5)
    }
    
    # 显示导出中通知
    showNotification("🔄 正在导出数据，请稍候...", type = "message", duration = NULL, id = "export_notif")
    selected_data <- all_data()[selected_rows, ]
    export_success <- export_selected_data(selected_data, output_path)
    
    # 移除导出中通知，展示结果
    removeNotification("export_notif")
    if (export_success) {
      # 更新导出历史（最多保留5条，按时间倒序）
      new_record <- list(
        time = format(Sys.time(), "%Y-%m-%d %H:%M:%S"),
        count = nrow(selected_data),
        path = output_path
      )
      current_history <- export_history()
      current_history <- c(list(new_record), current_history)
      if (length(current_history) > 5) {
        current_history <- current_history[1:5]
      }
      export_history(current_history)
      
      # 渲染导出成功信息
      output$export_result <- renderPrint({
        cat("✅ 导出成功！\n")
        cat("📊 选中记录数：", nrow(selected_data), "\n")
        cat("📁 文件保存路径：\n", output_path)
      })
      showNotification("✅ 数据导出成功！", type = "message", duration = 5)
    } else {
      # 渲染导出失败信息（带排查提示）
      output$export_result <- renderPrint({
        cat("❌ 导出失败！\n")
        cat("📌 请检查：1. 文件是否被占用 2. 路径权限是否足够 3. 文件名是否合法")
      })
      showNotification("❌ 数据导出失败！", type = "error", duration = 5)
    }
  })
  
  # 7. 渲染最近导出记录（按时间倒序显示，最多5条）
  output$export_history <- renderPrint({
    history_list <- export_history()
    if (length(history_list) == 0) {
      cat("📜 暂无导出记录")
    } else {
      cat("📜 最近导出记录（按时间倒序）：\n\n")
      for (i in 1:length(history_list)) {
        record <- history_list[[i]]
        cat(paste0("【", i, "】 时间：", record$time, "\n"))
        cat(paste0("    记录数：", record$count, "\n"))
        cat(paste0("    路径：", record$path, "\n\n"))
      }
    }
  })
}

#### 启动Shiny应用（本地运行，无网络传输，数据闭环，直接运行即可）
shinyApp(ui = ui, server = server)
test 