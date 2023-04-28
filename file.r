library(readxl)
library(dplyr)
library(tidyr)
library(lubridate)
library(stringr)
library(ggplot2)

# Function 1: Read in an Excel file
read_excel_file <- function(file_path) {
  library(readxl)
  data <- read_excel(file_path)
  return(data)
}

# ... (All the other existing functions you provided)

# Function 24: Analyze data
analyze_data <- function(data) {
  # Perform your desired analysis on the dataset here
  # Replace this line with your analysis code
  result <- "Replace this with your analysis result"
  return(result)
}


# Function 2: Sum all values in a column
sum_column <- function(data, column_name) {
  total_sum <- sum(data[[column_name]], na.rm = TRUE)
  return(total_sum)
}

# Function 3: Calculate the mean of a column
mean_column <- function(data, column_name) {
  column_mean <- mean(data[[column_name]], na.rm = TRUE)
  return(column_mean)
}

# Function 4: Calculate the median of a column
median_column <- function(data, column_name) {
  column_median <- median(data[[column_name]], na.rm = TRUE)
  return(column_median)
}

# Function 5: Count the number of rows in a dataset
count_rows <- function(data) {
  row_count <- nrow(data)
  return(row_count)
}

# Function 6: Count the number of columns in a dataset
count_columns <- function(data) {
  col_count <- ncol(data)
  return(col_count)
}

# Function 7: Filter dataset by a specific condition
filter_data <- function(data, column_name, condition) {
  library(dplyr)
  filtered_data <- data %>% filter(eval(parse(text = paste(column_name, condition))))
  return(filtered_data)
}

# Function 8: Group data by a specific column and calculate the sum
group_and_sum <- function(data, group_column, sum_column) {
  library(dplyr)
  grouped_data <- data %>% group_by(!!sym(group_column)) %>% summarise(sum = sum(!!sym(sum_column), na.rm = TRUE))
  return(grouped_data)
}

# Function 9: Group data by a specific column and calculate the mean
group_and_mean <- function(data, group_column, mean_column) {
  library(dplyr)
  grouped_data <- data %>% group_by(!!sym(group_column)) %>% summarise(mean = mean(!!sym(mean_column), na.rm = TRUE))
  return(grouped_data)
}

# Function 10: Group data by a specific column and calculate the median
group_and_median <- function(data, group_column, median_column) {
  library(dplyr)
  grouped_data <- data %>% group_by(!!sym(group_column)) %>% summarise(median = median(!!sym(median_column), na.rm = TRUE))
  return(grouped_data)
}

# Function 11: Convert a column of strings to lowercase
to_lowercase <- function(data, column_name) {
  data[[column_name]] <- tolower(data[[column_name]])
  return(data)
}

# Function 12: Convert a column of strings to uppercase
to_uppercase <- function(data, column_name) {
  data[[column_name]] <- toupper(data[[column_name]])
  return(data)
}

# Function 13: Extract a substring from a column of strings
extract_substring <- function(data, column_name, start, end) {
  data[[column_name]] <- substring(data[[column_name]], start, end)
  return(data)
}

# Function 14: Convert a column of strings to a date format
to_date_format <- function(data, column_name, format) {
  data[[column_name]] <- as.Date(data[[column_name]], format)
  return(data)
}

# Function 15: Convert a column of strings to a numeric format
to_numeric_format <- function(data, column_name) {
  data[[column_name]] <- as.numeric(data[[column_name]])
  return(data)
}

# Function 16: Replace missing values in a column with a specified value
replace_na <- function(data, column_name, value) {
  data[[column_name]][is.na(data[[column_name]])] <- value
  return(data)
}

# Function 17: Calculate the percentage of missing values in a column
percentage_missing_values <- function(data, column_name) {
  missing_values <- sum(is.na(data[[column_name]]))
  total_values <- length(data[[column_name]])
  percentage <- (missing_values / total_values) * 100
  return(percentage)
}

# Function 18: Remove rows with missing values in a specific column
remove_na_rows <- function(data, column_name) {
  data <- data[!is.na(data[[column_name]]), ]
  return(data)
}

# Function 19: Remove columns with a percentage of missing values greater than a specified threshold
remove_high_na_columns <- function(data, threshold) {
  column_percentages <- sapply(data, function(column) {
    percentage_missing_values(data, colnames(data)[column])
  })
  data <- data[, column_percentages <= threshold]
  return(data)
}

# Function 20: Merge two datasets by a common column
merge_data <- function(data1, data2, common_column) {
  library(dplyr)
  merged_data <- inner_join(data1, data2, by = common_column)
  return(merged_data)
}

# Function 21: Rename a column in a dataset
rename_column <- function(data, old_name, new_name) {
  library(dplyr)
  renamed_data <- data %>% rename(!!new_name := !!old_name)
  return(renamed_data)
}

# Function 22: Reorder columns in a dataset
reorder_columns <- function(data, new_order) {
  data <- data[, new_order]
  return(data)
}

# Function 23: Save dataset to a CSV file
write_csv_file <- function(data, file_path) {
  library(readr)
  write_csv(data, file_path)
}

