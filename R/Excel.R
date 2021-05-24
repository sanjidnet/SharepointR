int_to_excel_column <- function(input){
  excel_column <- ""
  while(input > 0){
    remainder <- (input - 1) %% 26;
    excel_column <- paste0(LETTERS[remainder + 1], excel_column)
    input <- floor((input - remainder) / 26)
  }
  return(excel_column)
}

#' Write excel file to a Sharepoint location using a template
#'
#' @param read_team_name name of team or drive
#' @param read_folder replace back slashes with forward slashes `/`. Feel free to keep spaces if present.
#' @param read_file include extension
#' @param write_team_name name of team or drive
#' @param write_folder replace back slashes with forward slashes `/`. Feel free to keep spaces if present.
#' @param write_file include extension
#' @param dta `data.frame` or `data.table`
#' @param preserve_class if true, this will recognize `currency`, `percentage`, `character` class.
#' These classes will be converted to the respective excel format thusly. The rest will be `General`
#' If false, everything is `General`.
#' @return
#' @export
#'
#' @examples
writeExcelToSharepoint <- function(read_team_name, read_folder, read_file, write_team_name, write_folder, write_file, dta, preserve_class = TRUE){
  # No point proceeding if there is no data to write
  if(!dim(dta)[1]){
    warning("THERE IS NOTHING TO WRITE")
    return()
  }
  write_drive_id <- getDriveId(team_name = write_team_name)
  # Web URL may be useful for a download URL
  write_folder_url <- httr::content(httr::GET(sprintf(
    "https://graph.microsoft.com/v1.0/drives/%s/root:/General/%s", write_drive_id,  write_folder = gsub("\\s", "%20", write_folder)),
    httr::add_headers("Authorization" = Sys.getenv("SHAREPOINT_TOKEN"))))$webUrl
  # write_folder_id <- getItemId(team_name = write_team_name, folder_path = write_folder)
  write_file_id <- getItemId(team_name = write_team_name, folder_path = write_folder, filename = write_file)
  # delete_message <- deleteSharepointItem(team_name = write_team_name, folder_name = write_folder, file_name = write_file)
  Sys.sleep(10);

  copySharepointItem(read_team_name = read_team_name, read_folder = read_folder, read_file = read_file, write_team_name = write_team_name, write_folder = write_folder, write_file = write_file)
  Sys.sleep(30);

  write_file_id <- getItemId(team_name = write_team_name, folder_path = write_folder, filename = write_file)
  sheetid <- httr::content(httr::GET(sprintf("https://graph.microsoft.com/v1.0/drives/%s/items/%s/workbook/worksheets/", write_drive_id, write_file_id), httr::add_headers("Authorization" = Sys.getenv("SHAREPOINT_TOKEN"))))$value[[1]]$id
  Sys.sleep(10);

  column_letter <- int_to_excel_column(dim(dta)[2])
  write_table_id <- httr::content(httr::POST(sprintf("https://graph.microsoft.com/v1.0/drives/%s/items/%s/workbook/worksheets/%s/tables/add", write_drive_id, write_file_id, sheetid),
                                 httr::add_headers("Authorization" = Sys.getenv("SHAREPOINT_TOKEN"), "Content-Type" = "application/json"),
                                 body = sprintf("{address: 'A1:%s%s', 'hasHeaders': true}", column_letter, (dim(dta)[1] + 1))))$id
  Sys.sleep(10);
  session_id <- httr::content(httr::POST(sprintf("https://graph.microsoft.com/v1.0/drives/%s/items/%s/workbook/createSession", write_drive_id, write_file_id),
                             httr::add_headers("Authorization" = Sys.getenv("SHAREPOINT_TOKEN")), body = "{'persistChanges': true}"))$id
  data.table::setDT(dta)
  message("Start writing: ", Sys.time())
  for(column_id in 1:dim(dta)[2]){
    column_name <- names(dta)[column_id]
    column_data <- dta[, column_id, with = FALSE]
    patch_response <- writeColumn(drive_id = write_drive_id, write_file_id, sheetid, session_id, dta = column_data, column_id = column_id, column_name = column_name, preserve_class)
    if(patch_response$status_code != 200){
      message("Second Try"); Sys.sleep(10)
      patch_response <- writeColumn(drive_id = write_drive_id, write_file_id, sheetid, session_id, dta = column_data, column_id = column_id, column_name = column_name, preserve_class)
    }
    # message(column_name, ": DONE WRITING AT: ", Sys.time())
  }
  message("Finish writing: ", Sys.time())
  return(write_folder_url)
}

writeColumn <- function(drive_id, write_file_id, sheetid, session_id, dta, column_id, column_name, preserve_class){
  temp <- jsonlite::toJSON(list(values = as.list(c(column_name, t(dta)))), pretty = TRUE, na = "null")
  iterator <- ceiling(as.numeric(object.size(temp)) / 512 / 1024 / 4) # conservative iterator; always >= 1
  entry_limit <- floor(dim(dta)[1] / iterator)

  if(preserve_class == TRUE){
    if(sapply(dta, class)[[1]] == "character"){
      formatting_request <- httr::PATCH(sprintf("https://graph.microsoft.com/v1.0/drives/%s/items/%s/workbook/worksheets/%s/range(address='%s2:%s%s')",
                                          drive_id, write_file_id, sheetid, int_to_excel_column(column_id), int_to_excel_column(column_id), (dim(dta)[1] + 1)),
                                  httr::add_headers("Authorization" = Sys.getenv("SHAREPOINT_TOKEN"), "Content-Type" = "application/json", "Workbook-Session-Id" = session_id),
                                  # numberFormat for Excel Text is "@". Yeah, I find it's funny too.
                                  body = jsonlite::toJSON(list(numberFormat = as.list(rep("@", dim(dta)[1]))), pretty = TRUE, na = "null"))
    }
    if(sapply(dta, class)[[1]] == "percentage"){
      formatting_request <- httr::PATCH(sprintf(
        "https://graph.microsoft.com/v1.0/drives/%s/items/%s/workbook/worksheets/%s/range(address='%s2:%s%s')",
        drive_id, write_file_id, sheetid, int_to_excel_column(column_id), int_to_excel_column(column_id), (dim(dta)[1] + 1)),
        httr::add_headers("Authorization" = Sys.getenv("SHAREPOINT_TOKEN"), "Content-Type" = "application/json", "Workbook-Session-Id" = session_id),
        body = jsonlite::toJSON(list(numberFormat = as.list(rep("0.0%", dim(dta)[1]))), pretty = TRUE, na = "null"))
    }
    if(sapply(dta, class)[[1]] == "currency"){
      formatting_request <- httr::PATCH(sprintf(
        "https://graph.microsoft.com/v1.0/drives/%s/items/%s/workbook/worksheets/%s/range(address='%s2:%s%s')",
        drive_id, write_file_id, sheetid, int_to_excel_column(column_id), int_to_excel_column(column_id), (dim(dta)[1] + 1)),
        httr::add_headers("Authorization" = Sys.getenv("SHAREPOINT_TOKEN"), "Content-Type" = "application/json", "Workbook-Session-Id" = session_id),
        body = jsonlite::toJSON(list(numberFormat = as.list(rep("$#,##0.00", dim(dta)[1]))), pretty = TRUE, na = "null"))
    }
  }

  # FIRST PATCH
  request <- httr::PATCH(sprintf("https://graph.microsoft.com/v1.0/drives/%s/items/%s/workbook/worksheets/%s/range(address='%s1:%s%s')",
                           drive_id, write_file_id, sheetid, int_to_excel_column(column_id), int_to_excel_column(column_id), (entry_limit + 1)),
                   httr::add_headers("Authorization" = Sys.getenv("SHAREPOINT_TOKEN"), "Content-Type" = "application/json", "Workbook-Session-Id" = session_id),
                   body = jsonlite::toJSON(list(values = as.list(c(column_name, t(dta[1:entry_limit])))), pretty = TRUE, na = "null"))

  # REMAINING PATCHES
  if(iterator == 1) return(request)

  remainder <- dim(dta)[1] %% iterator
  t <- 1
  while(t < iterator){
    start_at <- (t*entry_limit + 1)
    end_at <- (t*entry_limit + entry_limit)
    writer <- dta[start_at : end_at, ]

    #message("t: ", t, " -start_at: ", start_at, " -end_at: ", end_at)

    request <- httr::PATCH(sprintf("https://graph.microsoft.com/v1.0/drives/%s/items/%s/workbook/worksheets/%s/range(address='%s%s:%s%s')",
                             drive_id, write_file_id, sheetid, int_to_excel_column(column_id), (start_at + 1), int_to_excel_column(column_id), (end_at + 1)),
                     httr::add_headers("Authorization" = Sys.getenv("SHAREPOINT_TOKEN"), "Content-Type" = "application/json", "Workbook-Session-Id" = session_id),
                     body = jsonlite::toJSON(list(values = as.list(c(t(writer)))), pretty = TRUE, na = "null"))
    t <- t + 1
  }

  # LOOP ENDS message("t: ", t, " -remainder: ", remainder, " -rows: ", dim(dta)[1])
  if(remainder == 0) return(request)
  # FIX REMAINDER SCRIPT
  request <- httr::PATCH(sprintf("https://graph.microsoft.com/v1.0/drives/%s/items/%s/workbook/worksheets/%s/range(address='%s%s:%s%s')",
                           drive_id, write_file_id, sheetid, int_to_excel_column(column_id), (dim(dta)[1] - remainder + 2), int_to_excel_column(column_id), (dim(dta)[1] + 1)),
                   httr::add_headers("Authorization" = Sys.getenv("SHAREPOINT_TOKEN"), "Content-Type" = "application/json", "Workbook-Session-Id" = session_id),
                   body = jsonlite::toJSON(list(values = as.list(c(t(tail(dta, remainder))))), pretty = TRUE, na = "null"))
  return(request)
}
