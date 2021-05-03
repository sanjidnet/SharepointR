#' Title
#'
#' @param team_name acceptable values: `products`, `utils`
#'
#' @return
#' @export
#'
#' @examples
getDriveId <- function(team_name){
  site_id <- httr::content(httr::GET(sprintf("https://graph.microsoft.com/v1.0/sites/opdepot.sharepoint.com:/sites/%s",
    team_name), httr::add_headers("Authorization" = Sys.getenv("SHAREPOINT_TOKEN"))))$id
  return(httr::content(httr::GET(sprintf("https://graph.microsoft.com/v1.0/sites/%s/drive/", site_id), httr::add_headers("Authorization" = Sys.getenv("SHAREPOINT_TOKEN"))))$id)
}

#' Title
#'
#' @param team_name acceptable values: `products`, `utils`
#' @param folder_path replace back slashes with forward slashes `/`. Feel free to keep spaces if present.
#' @param filename include extension
#'
#' @return item identification to be used for reading / writing / copying / moving / deleting
#' @export
#'
#' @examples
getItemId <- function(team_name, folder_path, filename){

  driveid <- getDriveId(team_name)
  folder_path <- gsub("\\s", "%20", folder_path)

  item_id <- httr::content(httr::GET(sprintf("https://graph.microsoft.com/v1.0/drives/%s/root:/General/%s%s", driveid, folder_path, filename),
    httr::add_headers("Authorization" = Sys.getenv("SHAREPOINT_TOKEN"))))$id
  if(!is.null(item_id)) warning(sprintf("Issue with file %s, located at %s under team %s!", filename, folder_path, team_name))

  return(item_id)
}

