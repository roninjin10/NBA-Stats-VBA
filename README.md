The NBAStats class provides an interface for interacting with the NBA API.  

The scraping modules contain useful functions to scrape data using the NBAClass that I personally use in my spreadsheet.  They need to be refactored to be used in a different spreadsheet.

The following libraries must be checked in tools->References of the VBA window:
Microsoft XML, Microsoft Scripting Runtime, Microsoft Internet Controls

VBA-JSON must be available to parse the JSON from the API https://github.com/VBA-tools/VBA-JSON

Also included are modules that scrape a few pages on basketball reference and playtype data.  These need to be refactored as they contain memory leaks.