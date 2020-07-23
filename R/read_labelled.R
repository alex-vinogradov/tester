#' Imports Data Frame from Excel File with Variable and Value Labels
#' @param file Excel workbook with two sheets: for data and dictionary
#' @param data.sheet Name or number of sheet that contains data
#' @param vars.sheet Name or number of sheet that contains data dictionary
#' @param ... Additional parameters to readxl::read_excel() function
#' @examples
#' # read_labelled("example.xlsx")
#' # read_labelled("example.xlsx", "Data View", "Variable View")
#' @importFrom readxl read_excel
#' @importFrom jsonlite fromJSON
#' @importFrom jsonlite toJSON
#' @importFrom haven labelled
#' @importFrom tibble as_tibble
#' @importFrom writexl write_xlsx
#' @export
read_labelled <- function(file, data.sheet = 1, vars.sheet = 2, ...) {
  data <- read_excel(file, sheet = data.sheet, ...)
  if ("tbl" %in% class(vars.sheet)) {
    vars <- vars.sheet
  } else {
    vars <- read_excel(file, sheet = vars.sheet)
  }

  if (!"variable" %in% names(vars)) stop("There is no 'variable' column in dictionary", call. = FALSE)
  if (!"label" %in% names(vars)) stop("There is no 'label' column in dictionary", call. = FALSE)
  if (!"values" %in% names(vars)) stop("There is no 'values' column in dictionary", call. = FALSE)

  varnames <- names(data)

  for (var in vars$variable) {
    if (var %in% varnames) {
      varlab <- vars[vars$variable == var, "label", TRUE]
      if (is.na(varlab)) varlab <- NULL
      values <- vars[vars$variable == var, "values", TRUE]
      if (!is.na(values)) values <- unlist(fromJSON(values)) else values <- NULL
      data[[var]] <- labelled(data[[var]], values, varlab)
      attr(data[[var]], "format.spss") <- "F8.0"
      miss <- vars[vars$variable == var, "missing", TRUE]
      attr(data[[var]], "na_values") <- fromJSON(miss)
    } else {
      cat("Variable does not exist:", var, "\n")
    }
  }
  return(data)
}

#' Writes Labelled Data into Excel File with Variable and Value Labels
#' @param x Labelled data frame
#' @param file Excel workbook with two sheets: for data and dictionary
#' @examples
#' # write_labelled(pss, "example.xlsx")
#' @export
write_labelled <- function(x, file) {
  extract_labs <- function(variable) {
    label <- attr(variable, "label")
    values <- attr(variable, "labels")
    c(label = label, values = toJSON(lapply(values, "[")))
  }

  vars <- as_tibble(
    x = t(data.frame(lapply(x, extract_labs))),
    rownames = "variable"
  )

  write_xlsx(
    list("Data View" = x, "Variable View" = vars),
    path = file
  )

}

#' Recodes specified values to NA, also removes their labels
#' @param x Variable which values are to be recoded into NA
#' @param values Vector of values to recode
#' @examples
#' var <- c(1,1,2,2,2,7,8,9,0)
#' zap_values(var, 7:9)
#' zap_values(var, c(1,2,3))
#' @export
zap_values <- function(x, values) {
  x[x %in% values] <- NA
  l <- attr(x, "labels")
  attr(x, "labels") <- l[!l %in% values]
  return(x)
}

#' Recodes user missing values to NA, and removes their labels
#' @param x Variable which missing values will be recoded into NA
#' @examples
#' var <- c(1,1,2,2,2,7,8,9,0)
#' attr(var, "na_values") <- c(7,8,9)
#' zap_na(var)
#' @export
zap_na <- function(x) {
  values <- attr(x, "na_values")
  zap_values(x, values)
}
