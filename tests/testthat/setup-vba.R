skip_if_no_vba = function() {

  skip_on_cran()
  skip_on_ci()
  skip_on_os(c("mac", "linux", "solaris"))
  skip_if_not(system2("cscript", stderr = FALSE, stdout = NULL) == 0L)

}
