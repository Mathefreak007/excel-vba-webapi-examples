# VBAProject Documentation

## 1. Project Summary
This VBA project provides an interactive weather data tool integrated into an Excel workbook. It allows users to search for cities, list geographic location suggestions using the Open-Meteo geocoding API, and retrieve current weather data for a selected location. The retrieved information—including temperature, wind, weather condition, and a clickable map link—is displayed on the worksheet "WeatherDemo."

## 2. Modules and Their Roles
- **mWeatherGeo:** Implements all core functionalities related to querying geographic suggestions, retrieving current weather data from Open-Meteo APIs, rendering results to the worksheet, and providing utility helper functions for HTTP requests, JSON parsing, and data formatting.
- **Tabelle2 (Worksheet module):** Handles user interaction on the "WeatherDemo" worksheet, specifically trapping double-click events in the suggestion list area to trigger the retrieval and display of weather data.

## 3. Public Procedures

- `ListGeoSuggestions()` in **mWeatherGeo**  
  Reads the city name input from cell B1 on the "WeatherDemo" sheet, clears previous results, writes headers, and populates up to 10 location suggestions retrieved from the geocoding API.

- `GetWeatherData(Optional ByVal rowIndex As Long = 0)` in **mWeatherGeo**  
  Retrieves and displays current weather data for the location specified in the suggestion list at the given row (or the active cell's row if not provided). Shows detailed weather information below the suggestions on the worksheet, along with an OpenStreetMap link.

- `ShowCityOnMap()` in **mWeatherGeo**  
  Opens a web browser pointing to the OpenStreetMap location of the currently resolved latitude and longitude on the "WeatherDemo" sheet, if present.

- `Worksheet_BeforeDoubleClick(ByVal Target As Range, Cancel As Boolean)` in **Tabelle2**  
  Captures double-click events on the "WeatherDemo" worksheet in the range A5:I1048576 and invokes `GetWeatherData` for the selected row, preventing the default double-click action.

## 4. Dependency Analysis

- `Worksheet_BeforeDoubleClick` (in **Tabelle2**) calls `GetWeatherData` from **mWeatherGeo** when a user double-clicks a location row.
- `ListGeoSuggestions` calls `WriteSuggestionHeaders` and `WriteGeoSuggestions` internally to write data to the worksheet.
- `GetWeatherData` calls `GetCurrentWeather` to fetch weather data from the Open-Meteo API.
- `GetCurrentWeather` depends on `HttpGet` for fetching raw JSON data from the web.
- JSON parsing functions like `JsonPickObject`, `JsonPickString`, `JsonPickNumber`, and `CollectResultObjects` are used internally by `GetCurrentWeather` and `WriteGeoSuggestions` to handle API responses.
- Helpers such as `ToInvariant`, `FromInvariant`, `UrlEncode`, and `BuildLabel` are used across multiple procedures for consistent data formatting.
- `ShowCityOnMap` builds and opens a hyperlink to an external OpenStreetMap URL using coordinates found on the worksheet.
- Weather code translation functions `WeatherCodeToText` and `WeatherCodeToIcon` assist in converting numeric weather codes into human-readable text and Unicode icons.

## 5. Documentation

### Project Structure
- **Worksheet "WeatherDemo"**  
  The main user interface where users enter a city name (cell B1), receive a list of geographic location suggestions starting from row 4, and view current weather data displayed below the suggestions.

- **mWeatherGeo Module**  
  Contains all business logic: HTTP communication, JSON parsing, data retrieval, formatting, and writing to the worksheet.

- **Tabelle2 Worksheet Module**  
  Handles user interaction events on the "WeatherDemo" sheet, specifically double-clicking a location suggestion.

### Important Concepts
- **Geocoding API Integration:** The project uses the Open-Meteo geocoding API to translate city names into geographic location suggestions including latitude, longitude, timezone, population, and IDs.
- **Weather API Integration:** Uses Open-Meteo's current weather API to obtain up-to-date weather conditions by latitude and longitude.
- **Interactive UI:** Users type a city into B1 and click a button linked to `ListGeoSuggestions` to retrieve location options. Double-clicking a suggestion triggers detailed weather data retrieval.
- **Dynamic Data Writing:** Weather and location data are dynamically written to and cleared from specific worksheet ranges, with headers styled for readability.

### Key Parameters and Usage
- **City Input (B1, WeatherDemo sheet):** The city name to search for.
- **Max Suggestions Parameter:** Defaults to 10 when calling geocoding API, but can be adjusted in code.
- **RowIndex (GetWeatherData):** Specifies which suggestion row to fetch weather data for; defaults to the currently active row.
- **Localization:** Numeric values convert between invariant (dot as decimal) and locale-specific formats to handle different regional settings.
- **URLs:** Properly URL-encoded query strings ensure API queries work reliably.

### External Resources Accessed
- Open-Meteo Geocoding API: `https://geocoding-api.open-meteo.com/v1/search`
- Open-Meteo Weather API: `https://api.open-meteo.com/v1/forecast`
- OpenStreetMap URLs for interactive map viewing.

### Special Features
- Unicode weather icons provide a visually intuitive representation of weather conditions.
- Dynamic hyperlinks created in-cell open detailed weather location maps in an external browser.
- Custom JSON parsing functions allow processing of JSON without external libraries.
- Locale-aware number conversion handles international decimal separators gracefully.

### Typical Workflows
1. **List locations:** User enters a city name into `WeatherDemo!B1` and runs `ListGeoSuggestions` (via a button). Up to 10 location suggestions appear starting in A4.
2. **Select location:** User double-clicks one suggested location row, triggering `GetWeatherData` to fetch current weather data displayed beneath the list.
3. **View detailed map:** User runs `ShowCityOnMap` to open the location in OpenStreetMap in a browser.
4. **Repeat:** User can clear and enter a different city to start another cycle.

---

This project is easy to extend and customize by modifying the number of suggestions, adding more detailed weather parameters, or integrating other APIs. The modular design and comprehensive helper functions facilitate maintenance and adaptability.
