#' Retrieve Session Script Directory
#'
#' Retrieve the script directory created for this session.
#'
#' @return The path to the script directory.
#'
#' @export
script_dir = function() {
  normalizePath(get(".scriptdir", envir = topenv()), mustWork = TRUE)
}


#' Default Skeleton VBScript
#'
#' Retrieve the path to the default skeleton VBScript file.
#'
#' @param print If `TRUE`, print the file contents to the console.
#' @return The path to the default skeleton VBScript.
#'
#' @details The skeleton VBScript must use the following placeholders:
#' - `{macro_application}`: Placeholder for the macro application name,
#'   e.g., "Word", "Excel", etc.
#' - `{collection_type}`: Placeholder for the collection type
#'    containing the macro, e.g., "Documents" for Word,
#'    "Workbooks" for Excel, etc.
#' - `{macro_file}`: Placeholder for the path to the file containing
#'   the macro.
#' - `{macro_name}`: Placeholder for the name of the macro.
#' - `{macro_args}`: Placeholder for the macro arguments.
#'
#' @export
default_skeleton = function(print = FALSE) {
  skeleton_path =   system.file("scripts/skeleton.vbs",
    package = packageName(), mustWork = TRUE)
  if (print) {
    writeLines(readLines(skeleton_path))
  }
  skeleton_path
}

#' Build Macro Script
#'
#' Build a VBS script that calls the macro specified macro.
#'
#' @param macro_file The file containing the macro. Currently, ".xls*"
#'   and ".doc*" files are supported.
#' @param macro_name The name of the macro.
#' @param ... Arguments to pass to the macro. Type conversions in the
#'   resulting VBScript are defined based on the R class of each
#'   argument as follows:
#'   - numeric: cast to double using `Cdbl()`
#'   - integer: cast to integer using `CInt()`
#'   - character: cast to string using `CStr()`
#' @param .skeleton Path to the skeleton VBScript.
#'   See [default_skeleton()] for more information.
#' @return A character string containing VBScript text.
#'
#' @examples
#' example_file = system.file("examples", "data_importer.xlsm",
#'   package = "vbsgen")
#' macro_script(example_file, "importData",
#'   inputFile = NA_character_, targetSheet = NA_character_,
#'   targetRange = NA_character_, outputFile = NA_character_)
#'
#' @importFrom glue glue
#' @export
macro_script = function(macro_file, macro_name, ...,
  .skeleton = default_skeleton()) {
  # check path to file
  macro_file = normalizePath(macro_file, mustWork = TRUE)
  # argument handling
  arg_list = list(...)
  arg_names = names(arg_list)
  arg_positions = seq_along(arg_list) - 1L
  arg_types = unlist(lapply(arg_list, class))
  # script argument handling
  dim_args = paste(glue("Dim {arg_names}"), collapse = "\n")
  assign_args_list = glue("{arg_names} = args({arg_positions})")
  convert_args = ifelse(
    arg_types == "numeric", glue("CDbl({arg_names})"), ifelse(
    arg_types == "integer", glue("CInt({arg_names})"), ifelse(
    arg_types == "character", glue("CStr({arg_names})"), arg_names
  )))
  assign_args = paste(assign_args_list, collapse = "\n")
  macro_args = paste(convert_args, collapse = ", ")
  # identify application type
  macro_application = ifelse(
    grepl("\\.xls[a-z]*$", tolower(macro_file)), "Excel", ifelse(
    grepl("\\.doc[a-z]*$", tolower(macro_file)),"Word", NA_character_
  ))
  if (is.na(macro_application)) {
    stop("Could not identify application from file name.")
  }
  collection_type = switch(macro_application,
    "Excel" = "Workbooks",
    "Word" = "Documents"
  )
  # write vsb script
  vbs_skeleton = paste(readLines(.skeleton), collapse= "\n")
  glue(vbs_skeleton)
}


#' Macro As Function
#'
#' Generate an R function that acts as an interface to a VBA macro.
#'
#' @inheritParams macro_script
#' @return An R function that calls a VBScript via `system2()`.
#'
#' @examples
#' example_file = system.file("examples", "data_importer.xlsm",
#'   package = "vbsgen")
#' macro_function(example_file, "importData",
#'   dataFile = NA_character_, targetSheet = NA_character_,
#'   targetRange = NA_character_, outputFile = NA_character_)
#'
#' @importFrom glue glue
#' @export
macro_function = function(macro_file, macro_name, ...,
  .skeleton = default_skeleton()) {
  # generate script
  script_file = normalizePath(file.path(script_dir(),
    glue("{macro_name}.vbs")), mustWork = FALSE)
  writeLines(macro_script(macro_file, macro_name, ...,
    .skeleton = .skeleton), script_file)
  # format arguments for function
  arg_names = names(list(...))
  arg_string = paste(
    gsub("\\", "\\\\", shQuote(script_file), fixed = TRUE),
    paste(arg_names, collapse = ", "),
    sep = ", "
  )
  # create function
  function_code = glue("
    {{macro_name}} = function({{paste(arg_names, collapse = \", \")}}) {
      system2(\"cscript\", args = c({{arg_string}}))
    }
  ", .open = "{{", .close = "}}")  
  fun = eval(parse(text = function_code))
  fun
}
