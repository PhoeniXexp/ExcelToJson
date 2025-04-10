# ExcelToJson

ExcelToJson - Converts an Excel sheet to JSON

## Features
- Convert Excel (.xls, .xlsx) to JSON
- Customize output structure with path mapping
- Support for nested objects using dot notation
- Simple UI for schema editing and preview

## Usage
1. Download the latest version from [Releases page](https://github.com/PhoeniXexp/ExcelToJson/releases)
2. Run `ExcelToJson.exe`
3. Drag and drop an `.xls` or `.xlsx` file and save the JSON.
4. Edit the conversion schema in the text box:
   - Default generates flat structure
   - Use dot notation for nesting (e.g., `"address.city"`)
   - Rename fields by changing keys in the schema
5. Click "Save" to generate and save JSON

### Schema Example
```json
{
  "EmployeeID": "id",
  "Full Name": "personal.full_name",
  "Birth Year": "personal.birth_year",
  "Department": "department.name"
}
```

### Output Example
```json
{
  "id": "E1001",
  "personal": {
    "full_name": "John Smith",
    "birth_year": "1985"
  },
  "department": {
    "name": "Engineering"
  }
}
```

## Requirements
- [.NET 8 Runtime](https://dotnet.microsoft.com/download/dotnet/8.0) must be installed on your system

## Notes
- Only processes the first worksheet
- All values remain as strings (no type conversion) or null
- Schema paths are case-sensitive
- For complex schemas, validate JSON before saving

## Limitations
- No array support in schema (single objects only)