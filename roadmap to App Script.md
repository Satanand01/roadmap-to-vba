Perfect. I’ll prepare a detailed, daily topic-wise roadmap to learn Google Apps Script from scratch with a focus only on Google Sheets. It will include videos, written tutorials, and interactive exercises where possible—designed to fit about one hour of study per day.

I’ll let you know as soon as it’s ready.


# Google Apps Script (GAS) for Google Sheets – 15-Day Learning Roadmap

## Day 1: Introduction – What is Apps Script & Opening the Editor

* **Learning Objectives:** Understand what Apps Script is (a JavaScript-based **cloud scripting language** for extending Google Sheets). Learn to open the Apps Script editor bound to a Google Sheet. Get familiar with the editor interface and run a simple function.
* **Resources:** Google Developers – *Apps Script Fundamentals* codelab; Ben Collins’ GAS Beginner’s Guide; Official docs “What is Apps Script” and Spreadsheet service overview. (Also try a YouTube intro tutorial like *“Google Apps Script for Beginners”*).
* **Overview:** Google Apps Script is a **cloud-based development platform** that lets you write small programs to automate and extend Google Sheets and other Workspace apps. Its built-in *Spreadsheet service* can create/modify spreadsheets, read/write cell data, create custom menus, and more. The editor is online; in a Sheet go to **Extensions > Apps Script** to open it.
* **Exercises:**

  * Create a new Google Sheet. Open **Extensions > Apps Script** to launch the editor. Rename the project.
  * In the code editor, write a function (e.g. `function hello() { SpreadsheetApp.getActiveSheet().getRange("A1").setValue("Hello"); }`) and click ▶️ to run it (grant permission if prompted).
  * Inspect the **Execution Log** or *Logger.log()* output to verify the script ran (see Logging in Apps Script).

## Day 2: Macros & Custom Functions

* **Learning Objectives:** Learn about *macros* (recorded actions) and *custom functions* (user-defined formulas) in Sheets. Record a simple macro and view its script. Write and use a basic custom function in a cell.
* **Resources:** Codelab – *Macros and Custom Functions*; Google support on Sheets Macros; Official docs on custom functions. YouTube: *“Record macros in Google Sheets”*.
* **Key Points:** A **macro** is a recorded series of actions you can replay in Sheets. Recording a macro automatically creates Apps Script code. A **custom function** is a script you call like `=MYFUNC()` in a cell, similar to built-in functions. For example, you can write:

  ```js
  /** @customfunction */ 
  function DOUBLE(x) { return x*2; }
  ```

  then use `=DOUBLE(A1)` in the sheet.
* **Exercises:**

  * Record a macro via **Extensions > Macros > Record macro**, perform a simple task (e.g. bold a cell), then save. Open **Tools > Script editor** to see the generated code.
  * Edit the macro script: replace a hard-coded range with `getActiveCell()`, save, and re-run.
  * In the editor, write a custom function (e.g. `DOUBLE` above). Back in Sheets, type `=DOUBLE(5)` in a cell and ensure it returns 10.

## Day 3: JavaScript Basics Refresher

* **Learning Objectives:** Review key JavaScript concepts used in Apps Script: variables (`let`, `const`), functions, loops (`for`, `forEach`), arrays, and objects. Understand that Apps Script is **based on JavaScript**.
* **Resources:** Codecademy or freeCodeCamp JavaScript tutorials; MDN Web Docs on JS basics. The codelab note: *“Apps Script is based on JavaScript”*.
* **Overview:** Apps Script uses standard JS syntax. For example, declare arrays and loop through them to process sheet data. You don’t need deep JS knowledge but should know basics (functions, loops, arrays). The \[Ben Collins guide]\[46] mentions this: GAS lets you extend Sheets by writing small programs, *“automating repeatable tasks”*, etc..
* **Exercises:**

  * In Apps Script, write a simple `for` loop that logs numbers 1–5 with `Logger.log(i)`. Run and check logs.
  * Use an array: e.g. `let arr = [10,20,30]; arr.forEach(x => Logger.log(x*2));` to see output.
  * Write a function that takes an array and returns the sum or transforms it (run and verify).

## Day 4: Logging and Debugging

* **Learning Objectives:** Learn how to use `Logger.log()` (or `console.log()`) to debug scripts, and how to view the Execution log.
* **Resources:** Official Logging guide; Apps Script IDE guide on Logger. YouTube: *“Apps Script Logger.log tutorial”*.
* **Key Points:** During development, use the **Execution log** in the Apps Script editor (View > Logs). The Apps Script docs note: *“A basic approach to logging… use the built-in execution log”*. `Logger.log("message")` writes to this log in real time.
* **Exercises:**

  * Add `Logger.log("Script started");` at the top of one of your functions, run it, and check the logs (Menu ▶️ > Execution log).
  * Introduce a deliberate error (e.g. typo a variable), run the script and observe the error message. Then fix it.
  * (Optional) Use the “Debugger” in the editor: set a breakpoint and step through code to inspect variables.

## Day 5: Spreadsheet and Range Basics

* **Learning Objectives:** Use the **Spreadsheet service** to get the active spreadsheet/sheet, and to read/write cell data and ranges.
* **Resources:** Codelab – *Spreadsheets, Sheets, and Ranges* (Fundamentals #2); Ben Collins’ guide; Google Developers API reference for [SpreadsheetApp](https://developers.google.com/apps-script/reference/spreadsheet/spreadsheet-app). Example code by Ayush Raj: setting A1 text.
* **Overview:** The `SpreadsheetApp` class lets you access the spreadsheet. For example, `SpreadsheetApp.getActiveSpreadsheet().getActiveSheet()` gets the current sheet. Then methods like `.getRange("A1").setValue("Hello")` read/write cells. As shown in a tutorial: *“getRange("A1").setValue("Hello, World!")”*.
* **Exercises:**

  * In Apps Script, write a function to set cell A1 to `"Hello World"` (as above). Run it and check the sheet.
  * Write a function to read a cell value: e.g. `let val = sheet.getRange("B2").getValue(); Logger.log(val);`.
  * Copy a block of data: use `getRange("A2:B5").getValues()` to read a 2D array and then `setValues()` to paste it elsewhere. Verify the copy.

## Day 6: Looping and Bulk Data Operations

* **Learning Objectives:** Process multiple rows/columns efficiently using loops and array methods. Use `getValues()` and `setValues()` for batch updates.
* **Resources:** Codelab – *Working with Data* (Fundamentals #3); Ayush Raj tutorial (loop example). Medium or StackOverflow on iterating ranges.
* **Key Points:** To update many cells, fetch them into an array: `let data = sheet.getRange("A2:A10").getValues();`. Then loop: e.g., as shown in a guide, *“rowNumbers.forEach(row => { sheet.getRange(row,1).setValue('Processed'); });”*. After processing, you can write arrays back with `setValues()`.
* **Exercises:**

  * Write a script that loops through rows 2–10 in column A and writes “Done” in column B of each row (using a `for` loop or `forEach`).
  * Use `getValues()` to read a whole column into a JS array, modify the array (e.g. multiply every number by 2), and use `setValues()` to write it back to the sheet.
  * Sort data by code: read all rows into an array, use `array.sort()`, then clear the range and write sorted data back.

## Day 7: Custom Menus in Sheets

* **Learning Objectives:** Add a custom menu to the Sheet’s UI and bind menu items to script functions. Use the `Ui` service.
* **Resources:** Official guide *“Custom Menus in Google Workspace”*; YouTube tutorial on custom menus. Example from Codelab: using `ui.createMenu()` inside `onOpen()`.
* **Overview:** Apps Script can extend Sheets by adding new menu items that trigger your functions. For example, in an `onOpen()` function you can use `SpreadsheetApp.getUi().createMenu('My Menu').addItem('Say Hi','sayHi').addToUi();` to add a “My Menu” with a “Say Hi” command. Note: the script must be **bound to the sheet** and run in an `onOpen(e)` trigger.
* **Exercises:**

  * Write an `onOpen()` function that creates a custom menu “My Tools” with an item “Highlight” that calls a function `highlightSheet()`. Save and run `onOpen`, then reload the sheet to see the menu.
  * In `highlightSheet()`, use `sheet.getActiveRange().setBackground('yellow')` to color the selected cell. Test it via the menu.
  * Add another menu item, e.g. “Log Date”, that calls a function writing the current date/time into a new row.

## Day 8: Simple Triggers – onOpen, onEdit

* **Learning Objectives:** Learn about **simple triggers** such as `onOpen(e)` and `onEdit(e)`, which run automatically on events. Write scripts that respond to user actions in the sheet.
* **Resources:** Google Developers – *Simple Triggers* guide; Codelab examples. YouTube: *“Google Sheets onEdit trigger tutorial”*.
* **Key Points:** Triggers let scripts run automatically when events occur. For Sheets, common simple triggers are `onOpen(e)` (runs when a user opens the spreadsheet) and `onEdit(e)` (runs on each cell edit). For example, an `onEdit(e)` function can detect `e.range` and act on the edited cell. (Remember, simple triggers run without prompts but cannot access services needing authorization.)
* **Exercises:**

  * Implement an `onEdit(e)` function that sets a note on the edited cell with the timestamp (using `e.range.setNote(new Date())`). Test by editing any cell.
  * Write an `onOpen(e)` that shows a welcome message: e.g. `SpreadsheetApp.getUi().alert("Welcome!");`. Reload the sheet to trigger it.
  * (Optional) Note that edits made by scripts do NOT trigger `onEdit` again (to avoid loops). Try editing via script and observe this behavior.

## Day 9: Installable Triggers & Time-Driven Automation

* **Learning Objectives:** Use **installable triggers** to extend beyond simple triggers. Set up a time-driven (clock) trigger to run a script on a schedule.
* **Resources:** Official *Installable Triggers* guide; Google documentation on time-driven triggers; Medium/StackOverflow examples.
* **Overview:** Installable triggers (created via the script editor or programmatically) can call services that need authorization and include time-driven triggers. The docs note that a *“time-driven trigger… is similar to a cron job”* and can run as often as every minute or as infrequently as monthly. For example, you could auto-sort your sheet every morning or send a daily summary.
* **Exercises:**

  * In the Apps Script editor, go to **Triggers** (clock icon) and create a new trigger: select your function (e.g. `sortData`) and set “Time-driven”, e.g. to run daily at 9:00 AM.
  * Alternatively, write code using `ScriptApp.newTrigger('myFunction').timeBased().everyHours(1).create();` in a setup function and run it once.
  * As a test, write a function that appends a timestamp row, then let the time trigger run it after 1–5 minutes. Check the sheet.

## Day 10: Interactive UI – Alerts and Prompts

* **Learning Objectives:** Use the `Ui` service to display dialogs: alerts, prompts, and confirmations. Capture user input via scripts.
* **Resources:** Google *Dialogs and Sidebars* guide; YouTube *“Apps Script prompt alert”*. Example: using `SpreadsheetApp.getUi().prompt()` to ask the user for input.
* **Key Points:** Apps Script can show built-in dialogs. A **prompt** displays a message, an input box, and OK/Cancel buttons (similar to `window.prompt()` in browser). For example, `let result = ui.prompt("Enter name:", "", ui.ButtonSet.OK_CANCEL);` pauses the script until the user responds. The prompt dialog looks like this:【44†】. You can then use `result.getResponseText()` to get the input.
  &#x20;*A simple prompt dialog in Google Sheets, created with `SpreadsheetApp.getUi().prompt()`.*
* **Exercises:**

  * Write a function that, when run, shows `ui.alert("This is a message")`. Test it.
  * Create a menu item or button that calls a function using `prompt`. E.g., ask the user’s name and then show an alert greeting them by name.
  * Use `ui.confirm()` (yes/no) to branch logic (e.g. “Are you sure?” before deleting a row).

## Day 11: Fetching External Data (APIs)

* **Learning Objectives:** Learn to call external web APIs with `UrlFetchApp.fetch()` and process JSON data. Import that data into your sheet.
* **Resources:** Google *External APIs* guide; Medium example (using OpenWeatherMap API); StackOverflow examples.
* **Overview:** Apps Script can interact with any public API. The docs explain you can use the `UrlFetch` service to make HTTP requests. For example, you might fetch JSON from an open API. Once you have an `HTTPResponse`, call `getContentText()` and then `JSON.parse()` to turn it into an object. Then write the relevant data into your spreadsheet.
* **Exercises:**

  * Choose a simple public API (for example, [http://api.open-notify.org/iss-now.json](http://api.open-notify.org/iss-now.json) for ISS location, or a currency rate API). Write a script using `UrlFetchApp.fetch(url)`.
  * Parse the JSON response with `JSON.parse()`, extract a field (e.g. current price or timestamp), and write it into a cell.
  * (Advanced) Loop to fetch multiple items (e.g. last 5 days of weather) and append each as a new row.

## Day 12: Formatting and Data Presentation

* **Learning Objectives:** Apply formatting to sheets via script: fonts, colors, number formats, and conditional formatting. Optionally create charts from data.
* **Resources:** Codelab – *Data Formatting* (Fundamentals #4); Apps Script [Range](https://developers.google.com/apps-script/reference/spreadsheet/range) methods; Chart Service docs.
* **Key Points:** You can format ranges in Apps Script: e.g., `range.setFontWeight("bold")`, `range.setBackground("#ff0000")`, `range.setNumberFormat("$#,##0.00")`, etc. The Data Formatting codelab shows how to transform JSON into a nicely formatted sheet. You can also create charts with the Chart service (or use spreadsheet charts programmatically).
* **Exercises:**

  * Write a function to format your data range: bold the header row, color every other row, and format one column as currency or date.
  * (Optional) Use `SpreadsheetApp.newChart()` methods to create and insert a chart (e.g. a bar chart of some range).
  * Add conditional formatting via script: e.g., if a cell’s value is >100, make it red: use `sheet.getRange("A2:A").setConditionalFormatRules()`.

## Day 13: Project – Automating a Simple Workflow

* **Learning Objectives:** Apply combined skills to a mini-project: design a script that automates a repetitive task in your sheet. Examples: generating a report, managing data entries, etc.
* **Resources:** Review all previous docs. Look up related solutions (e.g., blogs on “Apps Script project examples” or StackOverflow).
* **Project Ideas:** Automate a sample workflow, such as:

  * **Expense Tracker:** On edit, auto-add timestamp to a “Last Updated” column; a custom menu to “Summarize Expenses” that totals each category.
  * **Data Consolidator:** Button or trigger that copies and merges data from multiple sheets into a master summary sheet.
  * **Reminder Script:** Weekly trigger that emails (or creates a popup) with a status of pending tasks (stay within Sheets by writing to a “Reminder” sheet).
* **Practice:** Sketch out the steps in plain English, then implement in Apps Script using menus, triggers, and data functions. Test thoroughly on sample data.

## Day 14: Project – Advanced Use and Cleanup

* **Learning Objectives:** Expand your project: add error handling, logging, and polish the user interface (menus, dialogs). Learn best practices (e.g. avoid hard-coding, optimize loops).
* **Resources:** Apps Script best practices blogs; StackOverflow for specific issues; Advanced section of Apps Script guides.
* **Tasks:**

  * Add try/catch around your code to handle unexpected data and log errors with `Logger.log()`.
  * If you built custom menus, add separators or icons as needed. Remove unused functions.
  * Document your code with comments. Use meaningful function and variable names.
* **Practice:** Review your scripts, refactor any repetitive code (e.g. use helper functions). Share the sheet with a colleague and test the automation under a different account to ensure triggers and permissions work.

## Day 15: Next Steps – Publishing and Further Learning

* **Learning Objectives:** Learn how to deploy scripts (as add-ons or web apps) and explore advanced Apps Script services. Plan continued learning resources.
* **Resources:** Google Developers – *Apps Script Dashboard and Deployment* docs; Google Workspace Developer YouTube channel; online forums (StackOverflow \[appscript tag], Google Apps Script Community).
* **Summary:** By now you’ve covered fundamentals: script editor, JavaScript basics, the Spreadsheet service (ranges, values, formatting), menus, triggers, and APIs. Continue practicing by building more complex add-ons or exploring other services (e.g. Gmail, Calendar) if interested. The [Apps Script samples](https://developers.google.com/apps-script/samples) and Codelabs playlists are excellent next steps.
* **Practice:** Try publishing your script as a **container-bound add-on** (via *Deploy > Test deployments*). Explore the **Apps Script API** or [clasp](https://github.com/google/clasp) for version control. Join the GAS community to learn tips and get help on advanced topics.

**Sources:** Official Google Developers documentation and codelabs, expert blogs (Ben Collins, Ayush Raj), and other trusted tutorials. Each day’s guidance is based on these sources to ensure up-to-date and accurate information.
