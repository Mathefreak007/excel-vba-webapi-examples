## 1. Project Summary

This VBA project provides routing functionality using the OpenRouteService (ORS) API integrated within Excel. It supports input as city names or latitude/longitude coordinates, performs geocoding via Open-Meteo, calculates routes between locations, and visualizes them on an interactive Leaflet-based HTML map. The project automates querying, route calculation, results output, and map generation/viewing directly from Excel.

## 2. Modules and Their Roles

- **mRouting:** Main module handling user input parsing, geocoding, routing API requests, route results processing, and output writing into the Excel worksheet.
- **modHelper:** Provides generic helper functions for HTTP requests, simple JSON parsing, URL encoding, locale conversion between strings and doubles, and worksheet management.
- **mRouteHtmlViewer:** Responsible for creating a local interactive HTML file to display the route on a Leaflet map and opening it in the default browser; builds the polyline and map markers from route coordinates.

## 3. Public Procedures

- `GetRoute()` in **mRouting**  
  Entry point sub that reads input from the worksheet, resolves start/destination locations via geocoding or direct coordinates, calls the ORS routing service, generates the HTML map, and outputs results back to Excel.

- `ORS_Route_GeoJSON(lat1 As Double, lon1 As Double, lat2 As Double, lon2 As Double, mode As String, ByRef dist As Double, ByRef dur As Double, ByRef coords As Variant) As Boolean` in **mRouting**  
  Requests a route from OpenRouteService API based on coordinates and travel mode, parses the GeoJSON response, extracts distance, duration, and coordinate array.

- `HttpGet(url As String, ByRef body As String) As Boolean` in **modHelper**  
  Executes a synchronous HTTP GET request to the specified URL and returns the response body as string if successful.

- `FindMatchingBrace(s As String, startPos As Long) As Long` in **modHelper**  
  Finds the position of matching closing brace `}` for an opening brace `{` in a JSON string starting from `startPos`.

- `FindMatchingBracket(s As String, startPos As Long) As Long` in **modHelper**  
  Finds the position of the matching closing bracket `]` for an opening bracket `[` starting from `startPos`.

- `JsonPickObject(json As String, key As String) As String` in **modHelper**  
  Extracts a JSON object `{ ... }` corresponding to a top-level key from a JSON string.

- `JsonPickString(json As String, key As String) As String` in **modHelper**  
  Extracts the string value for a specified key from a JSON string.

- `JsonPickNumber(obj As String, key As String) As String` in **modHelper**  
  Extracts the numeric value (as string) for a specified key from a JSON object string.

- `CollectResultObjects(json As String, maxN As Long) As Variant` in **modHelper**  
  Collects up to `maxN` JSON objects contained inside a top-level `results` array from a JSON response.

- `FromInvariant(s As String) As Double` in **modHelper**  
  Converts a string with invariant decimal point `.` to local system double value considering locale decimal separator.

- `ToInvariant(d As Double) As String` in **modHelper**  
  Converts a local double value to a string using invariant decimal point `.` for use in URLs/JSON.

- `UrlEncode(s As String) As String` in **modHelper**  
  Encodes a string minimally for URL use, encoding reserved characters.

- `EnsureSheet(name As String) As Worksheet` in **modHelper**  
  Ensures a worksheet with a specific name exists in the workbook; if missing, creates it.

- `SaveRouteHtml(coords As Variant, Optional filePath As String = "", Optional title As String = "Excel Route Viewer") As String` in **mRouteHtmlViewer**  
  Generates a local HTML file visualizing the route coordinates on a Leaflet map and returns the file path.

- `OpenRouteHtml(coords As Variant, Optional title As String = "Excel Route Viewer")` in **mRouteHtmlViewer**  
  Generates and immediately opens the route HTML map in the default web browser.

## 4. Dependency Analysis

- `GetRoute` (**mRouting**) is the main procedure; it:
  - Calls `ParseLatLon` to check if input is coordinate pair.
  - Calls `GeoResolveTop1` to geocode city names to coordinates.
  - Calls `ORS_Route_GeoJSON` to query ORS routing API.
  - Calls `SaveRouteHtml` (**mRouteHtmlViewer**) to create map file.
  - Calls `WriteRouteOutput` to write results into worksheet.

- `GeoResolveTop1` (**mRouting**) calls `HttpGet` (**modHelper**) to get geocoding data and uses JSON helpers `CollectResultObjects`, `JsonPickNumber`, and conversion helper `FromInvariant`.

- `ORS_Route_GeoJSON` (**mRouting**) calls `HttpGet` to get routing data; uses JSON helpers `JsonPickObject`, `JsonPickNumber`; calls parsing helpers `ExtractCoordinatesArray` and `ParseCoordinatesToArray` to parse GeoJSON coordinates.

- `SaveRouteHtml` (**mRouteHtmlViewer**) uses the public helper `ToInvariant` (**modHelper**) for formatting numbers in the HTML/JavaScript code.

- All JSON parsing relies on `modHelper` JSON utilities for minimal string-based JSON extraction.

## 5. Documentation

### Overview

This VBA project enables Excel users to calculate travel routes between two locations using online services and visualize the results interactively on a map. The user inputs either city names or latitude/longitude pairs into a dedicated worksheet named `"RouteDemo"`. The project then:

1. Parses and validates input locations.
2. If city names, geocodes them to coordinates using the Open-Meteo geocoding API.
3. Requests route directions from the OpenRouteService API for the chosen travel mode (e.g., driving, walking, cycling).
4. Extracts route distance, duration, and coordinates from the API response.
5. Produces a standalone HTML map file with the route drawn using LeafletJS.
6. Writes route details and a hyperlink to open the map into the Excel worksheet.

### Project Structure

- **Worksheet `"RouteDemo"`**:  
  Input cells:  
  - B1: Start location (city name or "lat,lon")  
  - B2: Destination location  
  - B3: Travel mode (optional; defaults to driving)  
  
  Output cells start from row 6 downward, showing resolved coordinates, distance, duration, and a clickable map link.

- **mRouting module**: Core logic for processing inputs and orchestrating API calls and output display.

- **modHelper module**: Low-level utilities for HTTP communication, JSON text extraction, locale-specific string to number conversions, and worksheet ensuring.

- **mRouteHtmlViewer module**: Responsible for mapping route coordinates into an HTML/Leaflet visualization, saving and optionally opening it.

### Important Concepts

- **Geocoding**: The project uses Open-Meteo’s simple geocoding API to resolve city names to geographic coordinates. Only the top 1 result is used.

- **Routing**: OpenRouteService is used for route calculations. A valid API key must replace the placeholder `YOUR_API_KEY` in `mRouting`.

- **Coordinates Format**: Latitude and longitude are handled consistently, but OpenRouteService returns coordinates in GeoJSON format `[lon, lat]`. This is converted internally to arrays `[lat, lon]` for Leaflet compatibility.

- **Travel Modes Supported**: `"driving"`, `"car"` → driving-car; `"cycling"`, `"bike"` → cycling-regular; `"walking"`, `"walk"` → foot-walking. Default: driving.

- **JSON Parsing**: The project uses purely string-based methods to extract JSON objects and key values, avoiding dependencies on external libraries.

- **Locale Handling**: Users running Excel with different regional settings can still correctly parse and format floating numbers for requests and responses.

- **Output Map**: The route is rendered in an HTML file saved locally (by default in the TEMP folder), showing the route polyline and start/destination markers.

### Parameters & Usage

- To run the routing, invoke `GetRoute` after entering input values in the `"RouteDemo"` sheet.
- API key for OpenRouteService must be set before usage (`ORS_API_KEY` constant).
- Start and destination can be either coordinates formatted as `"lat,lon"` (example: `48.137,11.575`) or city name strings.
- Travel mode in cell B3 is optional; supports recognized keywords.

### External Resources

- HTTP requests are made to:
  - Open-Meteo Geocoding API: `https://geocoding-api.open-meteo.com/v1/search`
  - OpenRouteService Directions API: `https://api.openrouteservice.org/v2/directions/`
- Generated Leaflet map scripts and tile layers are loaded via remote URLs:
  - Leaflet CSS and JS from `unpkg.com`
  - OpenStreetMap tile servers for rendering map tiles
- Uses Microsoft XMLHTTP (`MSXML2.ServerXMLHTTP.6.0`) for HTTP requests.

### Typical Workflow

1. Enter start and destination in cell B1 and B2, optionally travel mode in B3.
2. Run the macro `GetRoute`.
3. The project parses inputs, performs geocoding as needed.
4. Queries ORS routing service.
5. On success, writes resolved data, distance, duration to the sheet.
6. Generates a local HTML map for visualization; provides a clickable link in cell B15.
7. User can click the link to view the route interactively in their browser.

### Error Handling and Limitations

- Displays message boxes if inputs are missing or cannot be resolved.
- Does not handle multiple geocoding results beyond the top one.
- Requires an internet connection and valid ORS API key.
- All JSON parsing relies on simple string manipulations; complex API changes may break parsing.
- Leaflet map requires internet access to load tiles and scripts.

---

