#' Returns a character vector with the string patterns within the expunge
#' character vector removed.
#'
#' Uses stri_detect_regex.
#' @examples
#' strngs <- c("abc", "d", "E", "Fh")
#' expunge <- c("abc", "D", "E")
#' remove_strings(strngs, expunge, ignore_case = TRUE)
#' remove_strings(strngs, expunge, ignore_case = FALSE)
#' @param .str character vector
#' @param expunge character vector of strings to be removed
#' @param ignore_case logical indication whether or not to be case specific.
#' @import stringi
#' @export
remove_strings <- function(.str, expunge, ignore_case = FALSE) {
  if (ignore_case) {
    tmp_str <- tolower(.str)
    tmp_expunge <- tolower(expunge)
  } else {
    tmp_str <- .str
    tmp_expunge <- expunge
  }
  keep <- rep(TRUE, length(.str))
  for (exp_str in tmp_expunge) {
    keep <- !stri_detect_regex(tmp_str, exp_str) & keep
  }
  .str[keep]
}
#' Creates an Excel worksheet if it is needed.
#'
#' Takes a workbook object, a dataframe, the name of a worksheet, whether or
#' not the sheet is to have a header, and whether or not the sheet is to be
#' created if it does not exists, and either creates a worksheet or does based
#' on the values of the arguments and whether or not the sheet already exists.
#'
#' @return NULL
#' @param wb workbook object
#' @param m_df dataframe to receive formated worksheet
#' @param sheet character vector with name of worksheet
#' @param header logical vector of length one having TRUE if a header is to be
#' created in the worksheet.
#' @param create logical vector of length 1 that is TRUE if a worksheet is to
#'  be created if not present.
#' @import stringi
#' @import XLConnect
#' @export
create_sheet_if_needed <- function(wb, m_df, sheet, header, create) {
  if (!create) {
    if (!existsSheet(wb, sheet)) {
      stop(stri_c("Worksheet '", sheet,
                  "' was expected and not found."))
    }
  } else {
    if (!existsSheet(wb, sheet)) {
      createSheet(wb, sheet)
      writeWorksheet(wb, m_df, sheet, startRow = 1, startCol = 1,
                     header = header)
    }
  }
}
#' Adds formated data from a dataframe to a worksheet in a provided Excel file
#'
#' Takes a dataframe, an Excel file name, and a list of styles to be used one
#' the worksheet if the user provided function evalutes to TRUE for a cell
#' within the dataframe.
#'
#' The test function has to be written so that TRUE is assigned to all of the
#' cells where the formatting is wanted.
#'
#' As long as there is a least one cell that is to be formatted, we
#' set the cell styles. Remember if a cell has a value of TRUE for more
#' than one of the fmt_list items, that last fmt_list item specifications
#' will overwrite prior values.

#' If the sheet does not exist and \code{create} is TRUE, one is created using
#' \code{m_df}, otherwise we stop with an error.
#'
#' @param m_df dataframe to receive formated worksheet
#' @param excel_file character vector of length 1 with file name of Excel
#' workbook
#' @param sheet character vector with name of worksheet
#' @param header logical vector of length one having TRUE if a header is to be
#' created in the worksheet.
#' @param fmt_list list of format list used to format worksheet
#' @param create logical vector of length one having TRUE if a new worksheet
#' is to be created
#' @return Excel file name with formatted worksheet.
#'
#' @import stringi
#' @import XLConnect
#' @examples
#' library(stringi)
#' library(XLConnect)
#'
#' ## borders can be any of the following line types
#' names(XLC[stri_detect_regex(names(XLC), pattern = "^BORDER.")])
#' ## colors can be any of the following
#' names(XLC[stri_detect_regex(names(XLC), pattern = "^COLOR.")])
#' ## fill_pattern can be any of the following
#' names(XLC[stri_detect_regex(names(XLC), pattern = "^FILL.")])
#' my_f1 <- function(x) {x > 7}
#' fmt_lst_1 <- list(test = my_f1,
#'                 wrap = TRUE,
#'                 fill_pattern = as.integer(XLC$FILL.SOLID_FOREGROUND),
#'                 foreground_color = as.integer(XLC$COLOR.LAVENDER),
#'                 border = list(side = "all",
#'                               type = as.integer(XLC$BORDER.THICK),
#'                               color = XLC$COLOR.DARK_BLUE)
#'                 )
#' my_df <- data.frame(ColA = stri_c("Name_", 1:4),
#'                     ColB = c(4, 7, 8, 9),
#'                     ColC = c(100, 300, 3000, 132),
#'                     ColD = 1:4)
#' result <- add_formated_worksheet(my_df, "example_wkbk.xlsx",
#'                                  sheet = "my_test",
#'                                  header = TRUE,
#'                                  fmt_list = list(fmt_lst_1), create = TRUE)
#'
#' my_f2 <- function(x) {(x %% 3) == 0}
#' fmt_lst_2 <- list(test = my_f2,
#'                 wrap = TRUE,
#'                 fill_pattern = as.integer(XLC$FILL.SOLID_FOREGROUND),
#'                 foreground_color = as.integer(XLC$COLOR.TAN),
#'                 border = list(side = "all",
#'                               type = as.integer(XLC$BORDER.DOUBLE),
#'                               color = XLC$COLOR.GREY_80_PERCENT)
#'                 )
#' my_other_df <- data.frame(LETTERS = LETTERS[1:10],
#'                           NUMBERS = 1:10,
#'                           FRACTIONAL = 1 / (1:10))
#' result <- add_formated_worksheet(my_df, "example_wkbk.xlsx",
#'                                  sheet = "my_test",
#'                                  header = TRUE,
#'                                  fmt_list = list(fmt_lst_1, fmt_lst_2),
#'                                  create = TRUE)
#' @export
add_formated_worksheet <- function(m_df, excel_file, sheet = sheet,
                                   header = TRUE, fmt_list,
                                   create = TRUE) {
  ## Create the file if it does not exist
  if (!file.exists(excel_file)) {
    wb <- loadWorkbook(excel_file, create = TRUE)
    create <- TRUE
  } else {
    wb <- loadWorkbook(excel_file)
  }
  ## This offset is used to adjust row numbers when a header row exists in
  ## the Excel sheet
  if (header) {
    row_offset <- 1
  } else {
    row_offset <- 0
  }
  ## These two commands create a matrix of row numbers and column numbers
  ## respectively. Thus, the value of every cell in fmt_row has the row number
  ## of that cell and the value of every cell in fmt_col has the column number
  ## of that cell.
  fmt_row <-  matrix(data = rep(1:nrow(m_df), each = ncol(m_df)),
                     nrow = nrow(m_df),
                     ncol = ncol(m_df), byrow = TRUE) + row_offset
  fmt_col <- matrix(data = rep(1:ncol(m_df), each = nrow(m_df)),
                    nrow = nrow(m_df),
                    ncol = ncol(m_df))
  ## Step through each list item in fmt_list, each of which has the formating
  ## instructions corresponding to what is to be done to all cells that
  ## produce of value of TRUE in the test function associated with the list
  ## item.
  for (i in seq_along(fmt_list)) {
    fmt <- fmt_list[[i]]
    ## If the sheet does not exist and create is TRUE, one is created using
    ## m_df, otherwise we stop with an error.
    create_sheet_if_needed(wb, m_df, sheet, header, create)
    ## I only want to look at styles, so I am removing items in the list
    ## "fmt" that are not styles. (Earlier versions had more items than "test").
    styles <- remove_strings(names(fmt), expunge = c("test"))

    fmt_cellstyle <- createCellStyle(wb)
    for (style in styles) {
      switch(style,
             wrap ={
               setWrapText(fmt_cellstyle, wrap = fmt$wrap)
             },
             fill_pattern = {
               setFillPattern(fmt_cellstyle, fill = fmt$fill_pattern)
             },
             foreground_color = {
               setFillForegroundColor(fmt_cellstyle,
                                      color = fmt$foreground_color)
             },
             border = {
               setBorder(fmt_cellstyle, side = fmt$border$side,
                         type = fmt$border$type,
                         color = fmt$border$color)
             },
             {stop(stri_c("A style value of '", style, "' is not valid."))})
    }
    ## This is the hardest part to understand.
    ## fmt$test has to be written so that TRUE is assigned to all of the cells
    ## where the formatting is wanted. Once that is working, this routine
    ## makes two vectors with values of the row numbers and column numbers to
    ## receive the selected formatting.
    fmt_df <- data.frame(row = as.integer(fmt_row[fmt$test(m_df)]),
                         col = as.integer(fmt_col[fmt$test(m_df)]))
    fmt_df <- fmt_df[complete.cases(fmt_df), ]
    ## Currently, I am setting the column width to match the content width.
    ## This should be an option.
    setColumnWidth(wb, sheet, column = 1:ncol(m_df), width = -1)

    ## As long as there is a least one cell that is to be formatted, we
    ## set the cell styles. Remember if a cell has a value of TRUE for more
    ## than one of the fmt_list items, that last fmt_list item specifications
    ## will overwrite prior values.
    if (nrow(fmt_df) > 0) {
      setCellStyle(wb, sheet = sheet, row = fmt_df$row,
                   col = fmt_df$col, cellstyle = fmt_cellstyle)
    }
  }
  saveWorkbook(wb)
  excel_file
}
