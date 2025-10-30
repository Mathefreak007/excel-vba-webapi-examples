## 1. Project Summary

This VBA project provides a solution to analyze personal full names listed in an Excel worksheet using the external parser.name API. It automatically parses each full name into components such as first name, middle names, last name, and retrieves additional information including gender with confidence, country of origin with certainty, and salutation. The parsed data is then written back into the Excel sheet for user review.

## 2. Modules and Their Roles

- **modNameIntelligence:** Implements the main functionality to fetch and parse names via the parser.name API and populate results into an Excel worksheet.
- **modHelper:** Provides utility functions for HTTP requests, JSON parsing, URL encoding (UTF-8), and worksheet management to support the main module.

## 3. Public Procedures

- `AnalyseNames()` in **modNameIntelligence**  
  Reads full names from column A in the worksheet "NameDemo", calls the parser API for each name, parses results, and writes detailed components and metadata into columns B through J.

- `ParseNameAPI(fullName As String) As Variant` in **modNameIntelligence**  
  Sends an HTTP GET request to the parser.name API for the given full name, extracts relevant data from the JSON response, and returns an array with parsed name parts and related information.

- `HttpGet(url As String, ByRef body As String) As Boolean` in **modHelper**  
  Performs a synchronous HTTP GET request to the specified URL, retrieves the response body as UTF-8 decoded text, and returns success status.

- `JsonFirstDataBlock(json As String) As String` in **modHelper**  
  Extracts the first JSON object from the "data" array within the given JSON text.

- `JsonPickObject(json As String, key As String) As String` in **modHelper**  
  Extracts a JSON object associated with a specified key from a JSON string.

- `JsonPickString(obj As String, key As String) As String` in **modHelper**  
  Returns the string value for a given key within a JSON object string.

- `JsonPickNumber(obj As String, key As String) As String` in **modHelper**  
  Returns the numeric value (as string) for a given key within a JSON object string.

- `FindMatchingBrace(s As String, startPos As Long) As Long` in **modHelper**  
  Finds the position of the closing brace `}` that matches the opening brace `{` at `startPos` within a string.

- `FindMatchingBracket(s As String, startPos As Long) As Long` in **modHelper**  
  Finds the position of the closing bracket `]` that matches the opening bracket `[` at `startPos` within a string.

- `JsonPickArray(json As String, key As String) As String` in **modHelper**  
  Extracts the raw JSON array associated with a specified key from a JSON string.

- `CollectObjectsFromArray(arr As String, Optional maxN As Long = 50) As Variant` in **modHelper**  
  Parses and collects up to `maxN` JSON objects within a JSON array string and returns them as an array of strings.

- `UrlEncodeUTF8(s As String) As String` in **modHelper**  
  Encodes a Unicode string into a URL-encoded UTF-8 format suitable for HTTP requests.

- `EnsureSheet(name As String) As Worksheet` in **modHelper**  
  Returns the worksheet with the given `name`, creating it if it does not exist.

## 4. Dependency Analysis

- `AnalyseNames` (**modNameIntelligence**) is the main entry point and calls:  
  - `EnsureSheet` (**modHelper**) to get or create the results worksheet.  
  - `ParseNameAPI` (**modNameIntelligence**) for each full name.  
- `ParseNameAPI` calls:  
  - `UrlEncodeUTF8` (**modHelper**) to encode the name in the URL.  
  - `HttpGet` (**modHelper**) to retrieve JSON results from the API endpoint.  
  - Various JSON parsing functions in **modHelper** including `JsonFirstDataBlock`, `JsonPickObject`, `JsonPickString`, `JsonPickNumber`, `JsonPickArray`, and `CollectObjectsFromArray` to extract structured data from the API response.

The JSON helper functions depend on `FindMatchingBrace` and `FindMatchingBracket` for parsing structure tokens.

## 5. Documentation

### Project Structure

The project consists of two standard modules:

- **modNameIntelligence:** Contains the domain-specific logic to analyze names via the parser.name API and populate Excel.
- **modHelper:** Contains reusable utility functions that handle HTTP communication, URL encoding, JSON parsing, and worksheet handling.

### Important Concepts

- **API Integration:**  
  The project integrates with the parser.name API, which provides structured parsing of personal names along with demographic data such as gender and nationality.

- **JSON Parsing within VBA:**  
  Due to the lack of native JSON support, the project implements custom JSON parsing routines that extract objects, arrays, strings, and numbers from raw JSON strings by searching for braces and keys.

- **UTF-8 URL Encoding:**  
  Special care is taken to encode Unicode characters in the URL query string using an accurate UTF-8 encoding approach to ensure compatibility with international names.

- **Data Workflow in Excel:**  
  The input is read from a single column ("A") in the worksheet "NameDemo". Each full name is sent to the API, parsed, and multiple attributes are written to adjacent columns (B through J).

### Parameters and Usage

- **API Key:**  
  Replace `"YOUR-APIKEY-HERE"` in `modNameIntelligence` with your actual free API key from parser.name.

- **Input Data:**  
  The full names should be placed in column A of the "NameDemo" sheet starting from row 2 (assuming row 1 contains the header).

- **Output Data:**  
  The parsed components and metadata are output starting from column B:  
  - First Name  
  - Middle Name(s)  
  - Last Name  
  - Gender  
  - Gender Confidence (calculated as 1 - gender deviation)  
  - Country Code  
  - Country Name  
  - Country Certainty  
  - Salutation

### External Resources

- **Worksheets:**  
  The project relies primarily on the worksheet named "NameDemo". If this sheet does not exist, it is created automatically.

- **APIs and Internet Access:**  
  The project communicates over HTTP to the parser.name API endpoint at `https://api.parser.name/`. Reliable internet access is a prerequisite.

### Special Features

- **Custom JSON Parsing Without External Libraries:**  
  The project avoids dependencies on external JSON libraries by manually parsing JSON strings using string functions and character matching, making it portable within any standard VBA environment.

- **UTF-8 Safe URL Encoding:**  
  Ensures that names with special or non-ASCII characters are encoded correctly, avoiding errors in API calls.

- **Automatic Sheet Management:**  
  The code automatically creates the results worksheet if it is missing, facilitating ease of use.

### Typical Workflow

1. Populate column A of the "NameDemo" worksheet with a list of full names (one per row) starting at row 2.

2. Open the VBA editor, and run the `AnalyseNames` subroutine from **modNameIntelligence**.

3. The procedure will create or find the "NameDemo" sheet, clear (and write) headers, process each name via the API, parse the JSON results, and write the parsed components and metadata back into columns B to J.

4. After completion, a message box notifies the user that name analysis is complete.

5. Review the output data in the worksheet. Errors in parsing a particular name will be indicated by `#ERROR` in the first name column.

---

This documentation should facilitate understanding, usage, and further development of the VBA project for users and developers new to this codebase.