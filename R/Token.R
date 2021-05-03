
.onLoad <- function(libname, pkgname){
  ##https://medium.com/@anoopt/accessing-sharepoint-data-using-postman-sharepoint-rest-api-76b70630bcbf
  clientid <- Sys.getenv("SHAREPOINT_CLIENTID")
  clientsecret <- Sys.getenv("SHAREPOINT_CLIENTSECRET")
  realm <- Sys.getenv("SHAREPOINT_REALM") # also directory_id
  #principal <- "00000003-0000-0ff1-ce00-000000000000"
  #target <- "opdepot.sharepoint.com"

  if(!length(clientid)) warning("Client ID not found! Please follow instructions in README file.")

  access_url <- sprintf("https://login.microsoftonline.com/%s/oauth2/v2.0/token", realm)
  access_token <- httr::content(httr::POST(url = access_url, body = list(
    grant_type = "client_credentials", client_id = clientid, client_secret = clientsecret,
    scope = "https://graph.microsoft.com/.default")))$access_token
  Sys.setenv(SHAREPOINT_TOKEN = sprintf("Bearer %s", access_token))

  if(!length(access_token)) warning("Could not generate access token. Please follow instructions in README file.")


  # message("Sharepoint token set to system environment variable as `SHAREPOINT_TOKEN` ")
  # return(sprintf("Bearer %s", access_token))
}

