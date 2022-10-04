#package startup message
.onAttach <- function(libname, pkgname) {
  packageStartupMessage(paste0("glatosExtras ",
                               utils::packageVersion("glatosExtras")))
}
