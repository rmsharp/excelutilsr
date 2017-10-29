#' Creates an Excel workbook with worksheets.
#'
#' @param file filename of workbook to be created
#' @param df_list list of data frames to be added as worksheets to workbook
#' @param sheetnames character vector of worksheet names
#' @param create Specifies if the file should be created if it does not
#' already exist (default is FALSE). Note that create = TRUE has
#' no effect if the specified file exists, i.e. an existing file is
#' loaded and not being recreated if create = TRUE.
#' @import XLConnect
#' @export
create_wkbk <- function(file, df_list, sheetnames, create = TRUE) {
  if (length(df_list) != length(sheetnames))
    stop("Number of dataframes does not match number of worksheet names")

  if (file.exists(file) & create)
    file.remove(file)

  wkbk <- loadWorkbook(filename = file, create = create)
  for (i in seq_along(df_list)) {
    sheetname <- sheetnames[i]
    df <- df_list[[i]]
    createSheet(wkbk, sheetname)
    writeWorksheet(wkbk, df, sheetname, startRow = 1, startCol = 1,
                   header = TRUE)
    setColumnWidth(wkbk, sheetname, column = 1:ncol(df), width = -1)
  }
  saveWorkbook(wkbk)
  wkbk
}
#' Returns a character vector with the string patterns within the expunge
#' character vector removed.
#'
#' Uses stri_detect_regex.
#' @examples
#' strngs <- c("abc", "d", "E", "Fh")
#' expunge <- c("abc", "D", "E")
#' remove_strings(strngs, expunge, ignore_case = TRUE)
#' remove_strings(strngs, expunge, ignore_case = FALSE)
#' @param c_str character vector
#' @param expunge character vector of strings to be removed
#' @param ignore_case logical indication whether or not to be case specific.
#' @import stringi
#' @export
remove_strings <- function(c_str, expunge, ignore_case = FALSE) {
  if (ignore_case) {
    tmp_str <- tolower(c_str)
    tmp_expunge <- tolower(expunge)
  } else {
    tmp_str <- c_str
    tmp_expunge <- expunge
  }
  keep <- rep(TRUE, length(c_str))
  for (exp_str in tmp_expunge) {
    keep <- !stri_detect_regex(tmp_str, exp_str) & keep
  }
  c_str[keep]
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
#' @param m_df dataframe to receive formatted worksheet
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
#' Set the cell styles to be used in a workbook
#'
#' Cell styles are according to the format matrix provided (\code{fmt}.
#' Creates an anonymous cell style within the fmt_cellstyle object.
#'
#' @return cell style formats based on fmt
#' @param fmt format list(s) used to format the worksheet
#' @param wb workbook object
#' @import XLConnect
#' @export
set_cellstyle <- function(fmt, wb) {
  styles <- remove_strings(names(fmt), expunge = c("test"))

  fmt_cellstyle <- createCellStyle(wb)
  for (style in styles) {
    switch(style,
           wrap = {
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
  fmt_cellstyle
}
#' Add cell styles to a worksheet created on the fly for an existing workbook.
#'
#' Forms a named formatted worksheet based on argument values and places it in
#' the provided workbook..
#'
#' @param wb workbook object
#' @param m_df dataframe to receive formatted worksheet
#' @param sheet character vector with name of worksheet
#' @param fmt_list list of format list(s) used to format the worksheet
#' @param fmt_row a matrix of row numbers where the value of every cell
#' has the row number of that cell.
#' @param fmt_col a matrix of column numbers the value of every cell has
#' the column number of that cell.
#' @return \code{wb}
#'
#' @importFrom stats complete.cases
#' @import XLConnect
#' @export
add_cellstyles_to_wb <- function(wb, m_df, sheet, fmt_list, fmt_row, fmt_col) {
  ## Step through each list item in fmt_list, each of which has the formatting
  ## instructions corresponding to what is to be done to all cells that
  ## produce of value of TRUE in the test function associated with the list
  ## item.
  for (i in seq_along(fmt_list)) {
    fmt <- fmt_list[[i]]
    ## I only want to look at styles, so I am removing items in the list
    ## "fmt" that are not styles. (Earlier versions had more items than "test").
    fmt_cellstyle <- set_cellstyle(fmt, wb)
    ## This is the hardest part to understand.
    ## fmt$test has to be written so that TRUE is assigned to all of the cells
    ## where the formatting is wanted. Once that is working, this routine
    ## makes two vectors with values of the row numbers and column numbers to
    ## receive the selected formatting.
    fmt_df <- data.frame(row = as.integer(fmt_row[fmt$test(m_df)]),
                         col = as.integer(fmt_col[fmt$test(m_df)]))
    fmt_df <- fmt_df[stats::complete.cases(fmt_df), ]
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
  wb
}
#' Adds formatting to an existing Excel worksheet or to a worksheet added by
#' the function.
#'
#' Adds formatting to an existing worksheet based on values of a user provided
#' dataframe or adds a worksheet using all of the data in a dataframe and then
#' adds the indicated formatting to the newly added worksheet. If the Excel
#' workbook does not exist, it is created if the \code{create} flag is set to
#' \code{TRUE}.
#'
#' Takes a dataframe, an Excel file name, and a list of styles to be used on
#' the worksheet if the user provided function evalutes to \code{TRUE} for one
#' or more cells within the dataframe.
#'
#' The user must provide a test function written so that it evaluates to
#' \code{TRUE} for the cells where the indicated formatting is wanted.
#'
#' As long as there is at least one cell that is to be formatted, cell styles
#' are set. If a cell has a value of \code{TRUE} for more
#' than one of the \code{fmt_list} items, that last \code{fmt_list} item
#' specifications will overwrite prior values.

#' If the sheet does not exist and \code{create} is TRUE, one is created using
#' \code{m_df}, otherwise the function stops with an error.
#'
#' @param m_df dataframe to receive formatted worksheet
#' @param excel_file character vector of length 1 with file name of Excel
#' workbook
#' @param sheet character vector with name of worksheet
#' @param header logical vector of length one having TRUE if a header is in the
#' existing worksheet. A header is automatically created when the worksheet is
#' created by this routine.
#' @param fmt_list list of format list(s) used to format the worksheet
#' @param create logical vector of length one having \code{TRUE} if a new
#' worksheet is to be created
#' @return Excel file name with formatted worksheet.
#'
#' @import stringi
#' @import XLConnect
#' @examples
#' library(stringi)
#' library(XLConnect)
#'
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
#'                               color = as.integer(XLC$COLOR.DARK_BLUE))
#'                 )
#' my_df <- data.frame(ColA = stri_c("name_", 1:4),
#'                     ColB = c(4, 7, 8, 9),
#'                     ColC = c(100, 300, 3000, 132),
#'                     ColD = 1:4, stringsAsFactors = FALSE)
#' result <- add_formatted_worksheet(my_df, "example_wkbk.xlsx",
#'                                  sheet = "my_test",
#'                                  header = TRUE,
#'                                  fmt_list = list(fmt_lst_1), create = TRUE)
#' my_f2 <- function(x) {
#'   sapply(x, function(x) {
#'     if (all(is.numeric(x))) {
#'       (x %% 3) == 0
#'     } else {
#'       rep(FALSE, length(x))
#'     }
#'   })
#' }
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
#'                           FRACTIONAL = 1 / (1:10), stringsAsFactors = FALSE)
#' result <- add_formatted_worksheet(my_other_df, "example_wkbk.xlsx",
#'                                  sheet = "my_2nd_test",
#'                                  header = TRUE,
#'                                  fmt_list = list(fmt_lst_1, fmt_lst_2),
#'                                  create = TRUE)
#'
#' @export
add_formatted_worksheet <- function(m_df, excel_file, sheet = sheet,
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
  ## If the sheet does not exist and create is TRUE, one is created using
  ## m_df, otherwise we stop with an error.
  create_sheet_if_needed(wb, m_df, sheet, header, create)

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
  wb <- add_cellstyles_to_wb(wb, m_df, sheet, fmt_list, fmt_row, fmt_col)
  saveWorkbook(wb)
  excel_file
}
