#' Sharepoint Drive Id
#'
#' @param team_name name of team or drive
#'
#' @return
#' @export
getDriveId <- function(team_name){
  site_id <- httr::content(httr::GET(sprintf("https://graph.microsoft.com/v1.0/sites/opdepot.sharepoint.com:/sites/%s",
    team_name), httr::add_headers("Authorization" = Sys.getenv("SHAREPOINT_TOKEN"))))$id
  return(httr::content(httr::GET(sprintf("https://graph.microsoft.com/v1.0/sites/%s/drive/", site_id), httr::add_headers("Authorization" = Sys.getenv("SHAREPOINT_TOKEN"))))$id)
}

#' Sharepoint Item Id
#'
#' @param team_name name of team or drive
#' @param folder_path replace back slashes with forward slashes `/`. Feel free to keep spaces if present.
#' @param filename include extension, leave alone if you're looking for item id of a folder only
#'
#' @return item identification to be used for reading / writing / copying / moving / deleting
#' @export
getItemId <- function(team_name, folder_path, filename = ""){

  driveid <- getDriveId(team_name)
  folder_path <- gsub("\\s", "%20", folder_path)

  item_id <- httr::content(httr::GET(sprintf("https://graph.microsoft.com/v1.0/drives/%s/root:/General/%s%s", driveid, folder_path, filename),
    httr::add_headers("Authorization" = Sys.getenv("SHAREPOINT_TOKEN"))))$id
  if(is.null(item_id)) warning(sprintf("Issue with file %s, located at %s under team %s!", filename, folder_path, team_name))

  return(item_id)
}

#' List files in a folder
#'
#' @param team_name name of team or drive
#' @param folder_path replace back slashes with forward slashes `/`. Feel free to keep spaces if present.
#'
#' @return `data.table` containing `File` (name of the file / subfolder) and `LastModified` date
#'
#' @importFrom data.table ":="
#' @export
listFiles <- function(team_name, folder_path){
  read_drive_id <- getDriveId(team_name); read_folder_id <- getItemId(team_name, folder_path)
  i <- 0
  file_repo <- data.table::data.table()
  url_this <- sprintf("https://graph.microsoft.com/v1.0/drives/%s/items/%s/children?top=4000&&select=name,LastModifiedDateTime", read_drive_id, read_folder_id)
  while(i < 100){
    response <- httr::GET(url_this, httr::add_headers("Authorization" = Sys.getenv("SHAREPOINT_TOKEN")))
    file_repo <- rbind(file_repo, data.table::data.table(t(data.table::setDT(httr::content(response)$value))))
    message(i, " : ", url_this)
    if(is.null(httr::content(response)$`@odata.nextLink`)){
      message("NO MORE PAGES LEFT")
      break
    }
    url_this <- httr::content(response)$`@odata.nextLink`
    i <- i + 1
    #Sys.sleep(1)
  }
  data.table::setnames(file_repo, c("FILE_ID", "LastModified", "File")); file_repo[, FILE_ID := NULL]
  file_repo[, LastModified := as.character(LastModified)]; file_repo[, LastModified := as.POSIXct(LastModified, format = "%Y-%m-%dT%H:%M%OS", tz = "UTC")]
  attributes(file_repo$LastModified)$tzone <- "Pacific/Auckland"
  return(file_repo)

}

