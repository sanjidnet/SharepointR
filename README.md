
<!-- README.md is generated from README.Rmd. Please edit that file -->

# SharepointR

<!-- badges: start -->
<!-- badges: end -->

The goal of SharepointR is to make use of Microsoft Graph API to access
and modify Microsoft Sharepoint items.

## Installation

Passwords, access tokens, API keys are not stored here. So, installation
will require the following steps:

1.  Open R Console / RStudio and find out R_USER folder by executing

    `Sys.getenv("R_USER")`

2.  Create / modify `.Rprofile` file in R_USER folder and
    ensure the following line exists in the file:

    `readRenviron(paste0(Sys.getenv("R_USER"), "/.Renviron"))`

3.  Create / modify `.Renviron` file in R_USER folder and
    ensure the following line exists in the file:

    `SHAREPOINT_CLIENTID` = &lt;&gt;

    `SHAREPOINT_CLIENTSECRET` = &lt;&gt;

    `SHAREPOINT_REALM` = &lt;&gt;

    Get these from [portal.azure.com](portal.azure.com). More
    instructions in Sharepoint Token Resource below:

``` r
devtools::install_github("sanjidnet/SharepointR")
```

REFERENCE

-   [Renviron Resource 1](http://www.dartistics.com/renviron.html)

-   [Renviron Resource
    2](https://support.rstudio.com/hc/en-us/articles/360047157094-Managing-R-with-Rprofile-Renviron-Rprofile-site-Renviron-site-rsession-conf-and-repos-conf)

-   [Sharepoint Token
    Resource](https://anoopt.medium.com/access-sharepoint-data-using-postman-eec5965400f2)
