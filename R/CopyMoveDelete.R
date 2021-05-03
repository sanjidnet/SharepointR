#' Copy Sharepoint Item
#'
#' @param read_team_name acceptable values: `products`, `utils`
#' @param write_team_name acceptable values: `products`, `utils`
#' @param read_folder replace back slashes with forward slashes `/`. Feel free to keep spaces if present.
#' @param write_folder replace back slashes with forward slashes `/`. Feel free to keep spaces if present.
#' @param read_file include extension
#' @param write_file include extension
#'
#' @return
#' @export
#'
#' @examples
copySharepointItem <- function(read_team_name, read_folder, read_file, write_team_name, write_folder, write_file){
  parameters <- processTransferParameters(read_team_name, read_folder, read_file, write_team_name, write_folder, write_file)
  response <- httr::POST(sprintf("https://graph.microsoft.com/v1.0/drives/%s/items/%s/copy", parameters$read_drive_id, parameters$read_file_id),
       httr::add_headers("Authorization" = Sys.getenv("SHAREPOINT_TOKEN"), "Content-Type" = "application/json"), body = sprintf(
         "{ \"parentReference\": {\"driveId\": \"%s\", \"id\": \"%s\"}, \"name\": \"%s\"}", parameters$write_drive_id, parameters$write_folder_id, gsub("%20", " ", write_file)))
  if(!response$status_code == 202) warning("Copy Unsuccessful! Check Response.")
  return(response)
}

#' Move Sharepoint Item
#'
#' @param read_team_name acceptable values: `products`, `utils`
#' @param write_team_name acceptable values: `products`, `utils`
#' @param read_folder replace back slashes with forward slashes `/`. Feel free to keep spaces if present.
#' @param write_folder replace back slashes with forward slashes `/`. Feel free to keep spaces if present.
#' @param read_file include extension
#' @param write_file include extension
#'
#' @return
#' @export
#'
#' @examples
moveSharepointItem <- function(read_team_name, read_folder, read_file, write_team_name, write_folder, write_file){
  if(read_team_name == write_team_name){ # move possible only within the same drive / team
    parameters <- processTransferParameters(read_team_name, read_folder, read_file, write_team_name, write_folder, write_file)
    response <- httr::PATCH(sprintf("https://graph.microsoft.com/v1.0/drives/%s/items/%s", parameters$read_drive_id, parameters$read_file_id),
      httr::add_headers("Authorization" = Sys.getenv("SHAREPOINT_TOKEN"), "Content-Type" = "application/json", "Prefer" = "respond-async"), body = sprintf(
        "{ \"parentReference\": {\"driveId\": \"%s\", \"id\": \"%s\"}, \"name\": \"%s\"}", parameters$write_drive_id, parameters$write_folder_id, gsub("%20", " ", write_file)))
    if(!response$status_code == 200) warning("Move Unsuccessful! Check Response.")
    return(response)
  } else { # move between drives has to be done by copying - then deleting original
    copy_status <- copySharepointItem(read_team_name, read_folder, read_file, write_team_name, write_folder, write_file)
    if(copy_status$status_code == 202) delete_sharepoint_item(team_name = read_team_name, folder_name = read_folder, file_name = read_file)
    return(copy_status)
  }
}

#' Delete Sharepoint Item
#'
#' @param team_name acceptable values: `products`, `utils`
#' @param folder_name replace back slashes with forward slashes `/`. Feel free to keep spaces if present.
#' @param file_name include extension
#'
#' @return
#' @export
#'
#' @examples
deleteSharepointItem <- function(team_name, folder_name, file_name){
  drive_id <- getDriveId(team_name); item_id <- getItemId(team_name, folder_path = folder_name, filename = file_name)
  # delete only if source exists
  if(!is.null(item_id)) httr::DELETE(sprintf("https://graph.microsoft.com/v1.0/drives/%s/items/%s", drive_id, item_id), httr::add_headers("Authorization" = Sys.getenv("SHAREPOINT_TOKEN")))

}

processTransferParameters <- function(read_team_name, read_folder, read_file, write_team_name, write_folder, write_file){
  read_folder <- gsub("\\s", "%20", read_folder); write_folder <- gsub("\\s", "%20", write_folder); read_file <- gsub("\\s", "%20", read_file)
  read_drive_id <- getDriveId(read_team_name); write_drive_id <- getDriveId(write_team_name);

  read_file_id <- httr::content(httr::GET(sprintf("https://graph.microsoft.com/v1.0/drives/%s/root:/General/%s%s", read_drive_id, read_folder, read_file), httr::add_headers("Authorization" = Sys.getenv("SHAREPOINT_TOKEN"))))$id
  if(is.null(read_file_id)) warning("No file to transfer because source does not exist")
  write_folder_id <- httr::content(httr::GET(sprintf("https://graph.microsoft.com/v1.0/drives/%s/root:/General/%s", write_drive_id, write_folder), httr::add_headers("Authorization" = Sys.getenv("SHAREPOINT_TOKEN"))))$id

  write_file_id <- getItemId(team_name = write_team_name, folder_path = write_folder, filename = write_file)
  if(!is.null(write_file_id)){
    warning("Destination already exists and will be overwritted")
    delete_sharepoint_item(team_name = write_team_name, folder_name = write_folder, file_name = write_file)

  }

  return(list(read_drive_id = read_drive_id, read_file_id = read_file_id, write_drive_id = write_drive_id, write_folder_id =  write_folder_id))
}
