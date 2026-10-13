# SeleniumBot

C# console bot that pages through IMDb keyword/list search results with Selenium (Chrome), collects each
title's rating, vote count and list position, computes a combined "Point" score and writes the sorted
table (Title, Rating, Vote, Point) into an Excel file. `Selenium.xlsx` and `Copy of Selenium.xlsx` are
sample outputs from the original runs.

## How to run

Requires the .NET 8 SDK and Chrome.

```
dotnet build
dotnet run
```

- The first prompt ("Enter integer after page=") is used as the base search URL (`val2 + "&start=" + n`), so give a full IMDb search URL. It then reads stdin until EOF (Ctrl-D); any lines are treated as extra search URLs whose titles are excluded.
- It writes into `Selenium.xlsx` in the current folder; set `SELENIUMBOT_XLSX` to use another existing workbook.

## Status

- Ported from .NET Framework 4.7.2 to `net8.0` (SDK-style project, Selenium.WebDriver 4.51.0).
  Excel output now uses ClosedXML instead of Office interop (no Excel install needed).
- The build was verified; **the bot itself was not run against IMDb**. The XPath selectors target IMDb's
  2020 page layout (`//*[@id="main"]/div/div[3]/...`) and are very likely out of date, so scraping will probably
  return nothing until they are updated.
