#' @importFrom utils packageName
.onLoad = function(libname, pkgname) {
  assign(".scriptdir", tempfile(packageName()), envir = topenv())
  dir.create(get(".scriptdir", envir = topenv()))
}
