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
  skeleton_path = system.file("scripts/skeleton.vbs",
    package = packageName(), mustWork = TRUE)
  if (print) {
    writeLines(readLines(skeleton_path))
  }
  skeleton_path
}


#' Get VBS Casting Function
#'
#' Identify the VBS casting function to use for an R object.
#'
#' @param arg An R primitive.
#' @return The VBS casting function to use.
#'
#' @importFrom glue glue
#' @keywords internal
arg_spec = function(arg) {
  if (isTRUE(is.na(arg))) {
    stop("argument is NA.")
  }
  switch(class(arg),
    "numeric" = "CDbl",
    "integer" = "CInt",
    "character" = "CStr",
    "logical" = "CBool",
    stop(glue("No known VBS conversion for type \"{class(arg)}\"."))
  )
}


#' Get Application
#'
#' Identify the macro application.
#'
#' @param filename A filename.
#' @return The Application.
#'
#' @importFrom glue glue
#' @keywords internal
application_spec = function(filename) {
  if (grepl("\\.xls[a-z]*$", tolower(filename))) {
    "Excel"
  } else if (grepl("\\.doc[a-z]*$", tolower(filename))) {
    "Word"
  } else {
    stop(glue("Could not identify application from filename \"{filename}\"."))
  }
}


#' Get Collection Type
#'
#' Identify the collection type.
#'
#' @param application An application, e.g. "Word" or "Excel".
#' @return The collection type.
#'
#' @importFrom glue glue
#' @keywords internal
collection_spec = function(application) {
  if (application == "Excel") {
    "Workbooks"
  } else if (application == "Word") {
    "Documents"
  } else {
    stop(glue("Did not recognize application \"{application}\"."))
  }
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
#' \dontrun{
#' example_file = system.file("examples", "data_importer.xlsm",
#'   package = "vbar")
#' macro_script(example_file, "importData",
#'   inputFile = NA_character_, targetSheet = NA_character_,
#'   targetRange = NA_character_, outputFile = NA_character_)
#' }
#'
#' @importFrom glue glue glue_data
#' @export
macro_script = function(macro_file, macro_name, ...,
  .skeleton = default_skeleton()) {
  # check path to file
  macro_file = normalizePath(macro_file, winslash = "\\", mustWork = TRUE)
  stopifnot(length(macro_file) == 1L)
  stopifnot(length(macro_name) == 1L)
  # argument handling
  arg_list = list(...)
  arg_names = names(arg_list)
  # VBS command-line arguments are zero-indexed
  arg_positions = seq_along(arg_list) - 1L
  arg_prefixes = unlist(lapply(arg_list, arg_spec))
  convert_args = glue("{arg_prefixes}({arg_names})")
  # identify application type
  macro_application = application_spec(macro_file)
  collection_type = collection_spec(macro_application)
  # write vsb script
  vbs_skeleton = paste(readLines(.skeleton), collapse = "\n")
  glue_data(vbs_skeleton, .x = list(
    dim_args = paste(glue("Dim {arg_names}"), collapse = "\n"),
    assign_args = paste(glue("{arg_names} = args({arg_positions})"),
      collapse = "\n"),
    macro_application = macro_application,
    collection_type = collection_type,
    macro_file = macro_file,
    macro_name = macro_name,
    macro_args = paste(convert_args, collapse = ", ")
  ))
}


#' Macro As Function
#'
#' Generate an R function that acts as an interface to a VBA macro.
#'
#' @inheritParams macro_script
#' @return An R function that calls a VBScript via `system2()`.
#'
#' @examples
#' \dontrun{
#' example_file = system.file("examples", "data_importer.xlsm",
#'   package = "vbar")
#' macro_function(example_file, "importData",
#'   dataFile = NA_character_, targetSheet = NA_character_,
#'   targetRange = NA_character_, outputFile = NA_character_)
#' }
#' @importFrom glue glue
#' @export
macro_function = function(macro_file, macro_name, ...,
  .skeleton = default_skeleton()) {
  # check path to file
  macro_file = normalizePath(macro_file, mustWork = TRUE)
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
