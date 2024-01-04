test_that("macro script generation", {

  tf = tempfile(fileext = ".xlsx")
  file.create(tf)
  on.exit(unlink(tf), add = TRUE)
  tf = normalizePath(tf, mustWork = TRUE)

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


test_that("macro function", {

  skip_if_no_vba()
  skip_if_not(requireNamespace("readxl"))

  test_data = data.frame(x = seq_len(10) * 1.0, y = seq_len(10) * 2.0)
  dummy_file = normalizePath(tempfile(fileext = ".csv"), mustWork = FALSE)
  write.csv(test_data, dummy_file, row.names = FALSE)

  example_file = normalizePath(
    system.file("examples", "data_importer.xlsm", package = "vbar"),
    mustWork = TRUE
  )

  macro_fun = macro_function(example_file, "importData",
    dataFile = character(), targetSheet = character(),
    targetRange = character(), outputFile = character())

  output_file = normalizePath(tempfile(fileext = ".xlsm"), mustWork = FALSE)

  expect_identical(
    macro_fun(dummy_file, "first_sheet", "A1", output_file),
    0L
  )
  expect_identical(
    as.data.frame(readxl::read_excel(output_file, "first_sheet")),
    test_data
  )

})
