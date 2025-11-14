## 1. Project Summary

This VBA project enables users to search for images from two online image repositories: the Metropolitan Museum of Art collection (Met Museum) and NASA's Image & Video Library. Users enter a search term and select one or both sources on an Excel worksheet; the project then fetches a limited set of relevant images with metadata, displays thumbnails in the sheet, and allows users to preview a larger image on double-click.

## 2. Modules and Their Roles

- **Tabelle1 (Worksheet Code Module):** Handles the worksheet event to trigger a large preview image when the user double-clicks a result row.
- **mSearch (Standard Module):** Core module containing all logic for searching, retrieving, parsing JSON data from Met and NASA APIs, displaying results and thumbnails, and managing HTTP requests with JSON helpers.

## 3. Public Procedures

- `SearchVisuals()` in **mSearch**: Main entry procedure triggered (e.g., by a button) to start the search process based on user input and settings on the "VisualSearch" worksheet.
- `InsertPreviewImage(ws As Worksheet, rowIndex As Long)` in **mSearch**: Inserts a large preview image anchored on the worksheet, corresponding to the selected search result row.
- `HttpGetUtf8(url As String, ByRef body As String) As Boolean` in **mSearch**: Makes an HTTP GET request to the specified URL and returns the UTF-8 response body.
- `UrlEncodeUTF8(s As String) As String` in **mSearch**: Encodes a string in UTF-8 URL encoding format for safe HTTP requests.
- `JsonPickString(obj As String, key As String) As String` in **mSearch**: Extracts a string value by key from simple JSON text.
- `JsonPickNumber(obj As String, key As String) As String` in **mSearch**: Extracts a numeric value by key from simple JSON text.
- `FindMatchingBrace(s As String, startPos As Long) As Long` in **mSearch**: Finds the position of the matching closing brace `}` in JSON text from a given opening brace `{`.
- `FindMatchingBracket(s As String, startPos As Long) As Long` in **mSearch**: Finds the position of the matching closing bracket `]` in JSON text from a given opening bracket `[`.

## 4. Dependency Analysis

- `SearchVisuals` is the main entry point; it reads user inputs and calls:
  - `ClearResultArea` to clear old search results and thumbnails.
  - `ClearPreviewImage` to clear any existing preview image.
  - `FetchMetResults` if the Met Museum source is enabled.
  - `FetchNasaResults` if the NASA source is enabled.

- `FetchMetResults` calls:
  - `HttpGetUtf8` to retrieve JSON search results from the Met API.
  - `ExtractBetween` to parse JSON text.
  - `ShuffleArray` to randomize result order.
  - For each selected object ID, it calls `WriteMetObject`.

- `WriteMetObject` calls:
  - `HttpGetUtf8` to get detailed JSON data of a Met object.
  - JSON extraction functions like `JsonPickString`.
  - `InsertThumbnail` to add a thumbnail image into the worksheet.
  - Updates the output row index by reference.

- `FetchNasaResults` calls:
  - `HttpGetUtf8` to retrieve JSON search results from NASA API.
  - `ExtractArray`, `FindMatchingBrace`/`FindMatchingBracket`, and `ShuffleArray` for JSON parsing and randomization.
  - Calls `WriteNasaItem` per item.

- `WriteNasaItem` calls:
  - JSON extraction helpers (`JsonPickArrayObject`, `JsonPickString`).
  - `InsertThumbnail` to insert the image thumbnail.
  - Updates row index by reference.

- `InsertPreviewImage` is called by the worksheet double-click event handler (`Worksheet_BeforeDoubleClick` in **Tabelle1**) to display a larger preview based on the clicked row.

## 5. Documentation

### Project Structure and Main Concepts

- The project centers around an Excel worksheet named "VisualSearch". This worksheet hosts user input cells and the search results table starting from row 6.
- Users enter a search term in cell B1, and specify which sources to use in cells B2 (Met Museum) and B3 (NASA) by entering "Y" or any other value.
- The `SearchVisuals` subroutine, typically triggered by a button, validates input and executes searches on the indicated sources.
- Search results are displayed starting at row 6, columns A through G:
  - A: Source ("Met" or "NASA")
  - B: Title
  - C: Artist or author (Met) or "NASA"
  - D: Date or creation date
  - E: Description or department
  - F: Image URL (hidden for user reference)
  - G: Thumbnail image inserted as a Shape.

### Important Procedures and Parameters

- **SearchVisuals()**
  - No parameters; reads inputs directly from the "VisualSearch" worksheet.
  - Clears old search data and thumbnails.
  - Uses `FetchMetResults` and `FetchNasaResults` to fill results.
  - Adjusts output row pointer by reference to allow sequenced insertion.

- **FetchMetResults(term As String, ws As Worksheet, ByRef rowOut As Long)**
  - `term`: Search keyword.
  - `ws`: Worksheet to write results.
  - `rowOut`: Starting row number for output, updated inside.

- **FetchNasaResults(term As String, ws As Worksheet, ByRef rowOut As Long)**
  - Parameters same as Met function.

- The JSON in both APIs is parsed with basic string scanning methods, avoiding heavy JSON libraries for lightweight and portable code.

- Thumbnails are inserted as native Excel shapes anchoring in the worksheet, scaled to fit within a square cell column ("G"), with the row height adjusted accordingly.

- When the user double-clicks any cell in the results area (rows 6 to 1000, columns 1 to 7), the worksheet event triggers `InsertPreviewImage` to show a larger image preview anchored at cell K4.

### External Resources

- **APIs Used:**
  - The Met Museum Collection API:  
    - Search endpoint: `https://collectionapi.metmuseum.org/public/collection/v1/search`  
    - Object details endpoint: `https://collectionapi.metmuseum.org/public/collection/v1/objects/{objectId}`

  - NASA Image & Video Library API:  
    - Search endpoint: `https://images-api.nasa.gov/search`

- **Worksheets:**
  - `"VisualSearch"` worksheet used extensively for:
    - Inputs: B1 (search term), B2/B3 (source toggles).
    - Output: rows 6+ for displaying result records.
    - Preview area: cells K4:N4 defined for large image preview anchoring.

- **Shapes:**
  - Thumbnails: Shape names prefixed with `"Thumb_"` + row number.
  - Preview image: Shape named `"PreviewImage"`, unique and replaced on each preview.

### Special Features and Workflows

- **Search Workflow:**
  1. User enters a search term in B1.
  2. User sets "Y" for Met and/or NASA in B2 and/or B3.
  3. User triggers `SearchVisuals` (via button or macro).
  4. The module clears old data and fetches fresh results from enabled APIs.
  5. Results are randomized and limited to 5 per source by default.
  6. Data and thumbnails appear in the sheet; images are fetched from URLs.

- **Preview Workflow:**
  - Double-clicking any cell in a search result row (6 to 1000) triggers a large preview of that item's image in a fixed region (K4:N4).
  - This preview respects image aspect ratio and resizes accordingly.
  
- **Error Handling:**
  - HTTP failures and JSON parsing failures are handled silently by exiting early.
  - Image insertion errors are ignored to maintain robustness.

- **JSON Parsing:**
  - The project uses simple, lightweight string manipulation to extract JSON key-values and arrays, avoiding external references.
  - Functions like `JsonPickString`, `ExtractBetween`, and matching brace/bracket finders parse small JSON fragments.

- **Randomization:**
  - Returned results arrays from APIs are shuffled before display for variety in repeated searches.

### Typical Usage

- Open the workbook and navigate to the "VisualSearch" worksheet.
- Type a keyword in B1, e.g., "moon".
- Set B2 to "Y" to include Met Museum, B3 to "Y" for NASA (or one of them).
- Click the search button linked to `SearchVisuals` or run macro.
- Review results starting at row 6.
- Double-click on any part of a result row's range to open a larger preview image.
- Repeat with new terms or source options.

---

This completes the comprehensive documentation suitable for developers and users newly introduced to the project.