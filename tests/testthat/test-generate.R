test_that("macro script generation", {

  tf = normalizePath(tempfile(fileext = ".xlsx"), mustWork = FALSE)
  file.create(tf)
  on.exit(unlink(tf), add = TRUE)
  result = macro_script(tf, "fake_macro", a = 1L, b = 2.0, c = "three")
  result = gsub(tf, "testfile.xlsx", result, fixed = TRUE)
  expect_snapshot(result)

})


test_that("macro script helpers", {

  # argument casting
  expect_identical(arg_spec(1L), "CInt")
  expect_identical(arg_spec(1.0), "CDbl")
  expect_identical(arg_spec("one"), "CStr")
  expect_identical(arg_spec(FALSE), "CBool")
  expect_error(arg_spec(NULL))
  expect_error(arg_spec(NA))
  expect_error(arg_spec(data.frame()))
  # application detection
  expect_identical(application_spec("foo.xlsx"), "Excel")
  expect_identical(application_spec("foo.xls"), "Excel")
  expect_identical(application_spec("foo.docx"), "Word")
  expect_identical(application_spec("foo.doc"), "Word")
  expect_error(application_spec("foo.txt"))
  expect_error(application_spec(NA))
  expect_error(application_spec(NULL))
  # collection type
  expect_identical(collection_spec("Excel"), "Workbooks")
  expect_identical(collection_spec("Word"), "Documents")
  expect_error(collection_spec("Other"))
  expect_error(collection_spec(NA))
  expect_error(collection_spec(NULL))

})