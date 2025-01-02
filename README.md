# Uncommon VBScript Bugs and Errors

This repository demonstrates some less common, but potentially tricky, errors that can occur in VBScript programming.  VBScript's late-binding nature and limited error handling capabilities can lead to subtle bugs that are difficult to track down.

The `bug.vbs` file contains examples of these errors. The `bugSolution.vbs` file provides solutions and best practices to avoid these issues. 

## Bugs Covered:

* **Late Binding and Type Mismatches:**  Incorrect type handling due to VBScript's late binding.
* **Hidden Type Conversions:** Unexpected behavior from automatic type conversions.
* **Unhandled Exceptions:** Lack of robust error handling mechanisms.
* **Incorrect Object Handling:** Issues with COM object management.
* **Array Issues:** Problems with array access and dimensions.
* **Scope Issues:** Errors related to variable and function scope.