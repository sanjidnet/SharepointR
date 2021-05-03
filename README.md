
<!-- README.md is generated from README.Rmd. Please edit that file -->

# SharepointR

<!-- badges: start -->
<!-- badges: end -->

The goal of SharepointR is to make use of Microsoft Graph API to access
and modify Microsoft Sharepoint items.

## Installation

Passwords, access tokens, API keys are not stored here. So, installation
will require the following steps:

1.  Create / modify `.Rprofile` file in location `Sys.getenv("R_USER")`
    to ensure the following:

    `readRenviron(paste0(Sys.getenv("R_USER"), "/.Renviron"))`

2.  Create / modify `.Renviron` file in location `Sys.getenv("R_USER")`
    to ensure the following:

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
