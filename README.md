# Excel VBA Modules

This is an evolving repo of custom Excel VBA that I've put together and found useful in many circumstances,
so i figured I'd share.

#### Sections

[Steps to Enable/Access the VBA Editor](#steps-to-enableaccess-the-vba-editor)

[Steps to Install Custom VBA Formulas Manually via Copy-Paste](#steps-to-install-custom-vba-formulas-manually-via-copy-paste)

[Steps to Install Custom VBA Formulas via Import ( .bas  or. .cls )](#steps-to-install-custom-vba-formulas-via-import--bas--or-cls-)

#### Custom Formulas

[Sum By Color](#sumbycolor----add-as-module-)

[Sum Conditionaly Colored Cells](#sumconditioncolorcells----add-as-module-)

[Previous Sheet](#prevsheet----add-as-module-)

[Key Differences between SumByColor and SumConditionColorCells](#key-differences-between-sumbycolor-and-sumconditioncolorcells)


#### Custom Actions

[Add Sheet Change Action](#-addsheetchangeaction-----add-to-thisworkbook-)

[Copy Conditionally Formated Fill Color](#copyconditionalfillcolor---add-as-module-and-userform)

----------------------------------------------------------------------------------------------

## Steps to Enable/Access the VBA Editor
Here’s a quick guide to enable the VBA Editor in Excel. 
The process involves accessing it (it’s built into Excel but hidden by default) and ensuring it’s available for use.

1. **Open Excel:**
   - Launch Microsoft Excel with any workbook (new or existing).

2. **Use the Keyboard Shortcut:**
   - Press `Alt + F11`. This is the universal shortcut to open the VBA Editor directly.
   - The Visual Basic for Applications window will appear, showing the "Project" explorer on the left and a code window on the right.

3. **Alternative: Enable Developer Tab (Optional for Ribbon Access):**
   - If you prefer accessing the VBA Editor via the Excel ribbon (instead of the shortcut):
     - Go to `File > Options` (or `Excel Options` in older versions).
     - Click `Customize Ribbon`.
     - In the right pane, check the box for `Developer` under "Main Tabs."
     - Click `OK` to save.
   - Now, on the Excel ribbon, you’ll see the `Developer` tab.
   - In the `Developer` tab, click `Visual Basic` (or `VBA Editor`) to open it.

4. **Check Macro Settings (If Blocked):**
   - If the VBA Editor doesn’t open or macros seem disabled:
     - Go to `File > Options > Trust Center > Trust Center Settings`.
     - Select `Macro Settings`.
     - Choose `Enable VBA macros (not recommended; potentially unsafe code can run)` or `Enable all macros` for testing (adjust based on your security needs).
     - Click `OK` twice to exit the menus.
   - Try `Alt + F11` again.

5. **Start Using It:**
   - Once open, you can insert modules, paste code, or edit existing VBA projects (as described in the previous response about installing custom formulas).

** Troubleshooting **
- **Shortcut Not Working?** Ensure your keyboard’s `Alt` key isn’t remapped or disabled by other software.
- **No Developer Tab Option?** The `Alt + F11` shortcut works regardless, so you don’t *need* the Developer tab unless you want ribbon access.
- **Excel Version:** This works in all modern Excel versions (Windows and Mac), though the Mac interface might look slightly different.

### Steps to Install Custom VBA Formulas Manually via Copy-Paste

1. **Open the VBA Editor:**
   - In Excel, press `Alt + F11` to open the Visual Basic for Applications (VBA) editor.

2. **Insert a Module:**
   - In the VBA editor, right-click on your workbook name in the "Project" window (usually on the left).
   - Select `Insert > Module`. This creates a new module (e.g., `Module1`) where you’ll place your code.
   - For workbook-specific code (like `ThisWorkbook` events), use the `ThisWorkbook` object instead of a module.

3. **Paste the Code:**
   - Copy the VBA code for your custom function (e.g., `SumByColor`, `PrevSheet`, etc.).
   - Paste it into the module or `ThisWorkbook` (depending on the code type: functions go in modules, event handlers go in `ThisWorkbook`).

4. **Save the Workbook:**
   - Close the VBA editor (`Alt + Q` or the red X).
   - Save your workbook in a macro-enabled format: `File > Save As > Excel Macro-Enabled Workbook (*.xlsm)`.

5. **Enable Macros:**
   - When you reopen the workbook, Excel may prompt you to enable macros. Click `Enable Macros` or `Enable Content` to allow the VBA code to run.

6. **Use the Function:**
   - In any cell, type your custom function like a built-in one, e.g., `=SumByColor(A1:A10, B1)` or `=PrevSheet(A1)`.
   - If it’s working, it’ll calculate based on the code’s logic.

#### Notes
- **Module vs. ThisWorkbook:**
  - Use a **module** for user-defined functions (UDFs) like `SumByColor` or `PrevSheet`.
  - Use **`ThisWorkbook`** for event-driven code (e.g., `Workbook_SheetActivate`).
- **Security:** Ensure macros are enabled each time you open the file, or place the code in a trusted location.
- **Testing:** Test with simple data first to confirm the function behaves as expected.

### Steps to Install Custom VBA Formulas via Import ( .bas  or. .cls )

- Here’s a brief guide on how to import a custom VBA `.bas` (module) or `.cls` (class module) file into Excel to use its code:

#### Steps to Import a `.bas` or `.cls` File

1. **Open the VBA Editor:**
   - In Excel, press `Alt + F11` to open the VBA Editor.

2. **Locate Your Project:**
   - In the "Project" window (usually on the left), find your workbook’s name (e.g., `VBAProject (Book1)`).

3. **Import the File:**
   - Go to `File > Import File` in the VBA Editor menu (or right-click the project name and select `Import File`).
   - Browse to the location of your `.bas` (standard module) or `.cls` (class module) file.
   - Select the file and click `Open`.
   - The imported module or class will appear under your project (e.g., `Module1` for a `.bas` file or the class name for a `.cls` file).

4. **Verify the Code:**
   - Double-click the imported item in the Project window to view its code in the editor. Ensure it looks correct.

5. **Save the Workbook:**
   - Close the VBA Editor (`Alt + Q` or the red X).
   - Save your workbook as an Excel Macro-Enabled Workbook (`.xlsm`) via `File > Save As > Excel Macro-Enabled Workbook`.

6. **Test the Code:**
   - If it’s a function (e.g., from a `.bas` file), use it in a cell (e.g., `=YourFunctionName()`).
   - If it’s a class (from a `.cls` file), it’s typically used in other VBA code rather than directly in worksheets.

#### Notes
- **`.bas` Files:** Contain standard VBA code (e.g., functions like `SumByColor` or `PrevSheet`).
- **`.cls` Files:** Contain class definitions (e.g., `ThisWorkbook` or custom objects).
- **Duplicates:** If a module/class with the same name already exists, VBA may rename the imported one (e.g., `Module1_1`).
- **Macros Enabled:** Ensure macros are enabled when reopening the workbook (`Enable Content` prompt).

----------------------------------------------------------------------------------------------

## Custom Formulas

### `SumByColor - * Add as Module *

#### ** Purpose: ** 
Sums the values of cells in a specified range that match the background color of a reference cell.

#### ** How it works: **
- **Inputs:**
  - `SumRange`: The range of cells where values will be summed.
  - `SumColor`: A single cell whose background color (Interior.ColorIndex) is used as the reference color.
- **Logic:**
  1. Retrieves the `ColorIndex` (a numeric value representing the color) of the `SumColor` cell.
  2. Loops through each cell in `SumRange`.
  3. If a cell’s background color matches the `SumColor`’s `ColorIndex`, its value is added to `TotalSum`.
  4. Returns the total sum of matching cells.
  
- **Key Feature:** It only considers the **interior (background) color** of cells, not conditional formatting rules.
- **Volatility:** Marked with `Application.Volatile`, meaning it recalculates whenever any change occurs in the workbook.

#### ** Example Use Case: **
- You have a range `A1:A10` with numbers, and some cells are manually colored yellow. You use `=SumByColor(A1:A10, B1)` where `B1` is a yellow cell. The function sums all values in `A1:A10` where the background color is yellow.

#### ** Limitations: **
- Ignores conditional formatting (only looks at manually applied or static background colors).
- Requires an exact `ColorIndex` match.

----------------------------------------------------------------------------------------------

### `SumConditionColorCells - * Add as Module *

#### ** Purpose: ** 
Sums the values of cells in a range based on a color applied via **conditional formatting**, provided the condition evaluates to `True`.

#### ** How it works: **
- **Inputs:**
  - `CellsRange`: The range of cells to evaluate and sum.
  - `ColorRng`: A single cell whose background color (Interior.ColorIndex) is used to identify matching conditional formatting rules.
- **Logic:**
  1. Checks if any conditional formatting rule in `CellsRange` uses the same `ColorIndex` as `ColorRng`.
     - Loops through all conditional formatting rules in `CellsRange`.
     - If a match is found, sets `Bambo = True` and notes the rule’s index (`CF1`).
  2. If no matching color is found in the conditional formatting, returns `"NO-COLOR"`.
  3. If a match is found:
     - Loops through each cell in `CellsRange`.
     - Extracts the conditional formatting formula (e.g., `=A1>10`) for the matched rule.
     - Converts the formula between `A1` and `R1C1` notation to evaluate it relative to each cell.
     - If the formula evaluates to `True` for a cell, adds that cell’s value to `CF2`.
     - Returns the total sum (`CF2`).
	 
- **Key Feature:** Focuses on **conditional formatting colors** and evaluates the associated conditions.
- **Volatility:** Also marked with `Application.Volatile`, so it recalculates on workbook changes.

#### ** Example Use Case: **
- You have a range `A1:A10` with a conditional formatting rule like “if value > 5, color yellow.” You use `=SumConditionColorCells(A1:A10, B1)` where `B1` is a yellow cell. The function sums values in `A1:A10` where the condition `>5` is true and the conditional formatting applies yellow.

#### ** Limitations: **
- Only works with conditional formatting colors, not manual background colors.
- Requires the conditional formatting rule’s color to match `ColorRng`’s color.
- More complex and potentially slower due to formula evaluation.

----------------------------------------------------------------------------------------------

### Key Differences between SumByColor and SumConditionColorCells

| Feature                  | `SumByColor` (Module1)                  | `SumConditionColorCells` (Module2)             |
|--------------------------|-----------------------------------------|-----------------------------------------------|
| **Color Source**         | Manual/static background color          | Conditional formatting color                  |
| **Condition Evaluation** | None (just matches color)               | Evaluates conditional formatting formula      |
| **Output**               | Numeric sum                             | Numeric sum or `"NO-COLOR"` if no match       |
| **Use Case**             | Sum cells with a specific manual color  | Sum cells meeting a condition with a color    |
| **Complexity**           | Simple and straightforward              | More complex (handles formulas and conditions)|
| **Performance**          | Faster (simple color check)             | Slower (evaluates formulas for each cell)     |


#### Practical Implications
- Use `SumByColor` if you’re working with **manually colored cells** (e.g., someone highlighted cells in red using the fill tool) and just want to sum based on that color.
- Use `SumConditionColorCells` if you’re dealing with **conditional formatting** (e.g., cells turn green when a value exceeds a threshold) and need to sum based on both the color and the condition being met.

#### Example Scenario
- **Data:** Range `A1:A5` contains `[3, 7, 2, 8, 4]`.
- **Setup:**
  - `A2` and `A4` (7 and 8) are manually colored yellow.
  - Conditional formatting rule: “If value > 5, color yellow” applies to `A1:A5`, so `A2` (7) and `A4` (8) are also yellow via formatting.
  - `B1` is a yellow cell (manual color matching the conditional formatting color).
- **Results:**
  - `=SumByColor(A1:A5, B1)` → `15` (sums 7 + 8, based on manual yellow color).
  - `=SumConditionColorCells(A1:A5, B1)` → `15` (sums 7 + 8, based on the condition `>5` being true and yellow applied via conditional formatting).

In this case, the results match because the manually colored cells align with the conditional formatting. But if the manual colors and conditional formatting rules didn’t align, the outputs would differ.

----------------------------------------------------------------------------------------------

### `PrevSheet - * Add as Module *
**Purpose:** Returns a range object from the worksheet immediately preceding the one containing the input range. If there’s no previous sheet (i.e., the current sheet is the first one), it falls back to the same range on the current sheet.

#### ** How It Works: **
1. **Input Parameter:**
   - `rng As Range`: The function takes a range (e.g., `A1`, `B2:C5`) as input. This range is typically from the sheet where the function is called.

2. **Volatility:**
   - `Application.Volatile`: This line makes the function recalculate whenever any change occurs in the workbook, not just when its direct inputs change. This is useful for dynamic sheet references that might shift due to adding, deleting, or reordering sheets.

3. **Variable Setup:**
   - `Dim ws As Worksheet`: Declares a variable to hold a worksheet object.
   - `Set ws = rng.Parent`: Assigns `ws` to the worksheet containing the input range (`rng.Parent` is the sheet that "owns" the range).

4. **Logic:**
   - `If ws.Index > 1 Then`: Checks the position of the current worksheet in the workbook.
     - `ws.Index` is the numeric position of the sheet (e.g., 1 for the first sheet, 2 for the second, etc.).
     - If `ws.Index > 1`, the current sheet isn’t the first one, so there’s a previous sheet to reference.
   - `Set PrevSheet = ws.Parent.Worksheets(ws.Index - 1).Range(rng.Address)`:
     - `ws.Parent` is the workbook containing the sheet.
     - `Worksheets(ws.Index - 1)` accesses the previous sheet (e.g., if current sheet is index 2, it gets index 1).
     - `Range(rng.Address)` takes the address of the input range (e.g., `"A1"`, `"B2:C5"`) and applies it to the previous sheet, returning that range.
   - `Else`:
     - If `ws.Index = 1` (the current sheet is the first one), there’s no previous sheet.
     - `Set PrevSheet = ws.Range(rng.Address)`: Returns the same range from the current sheet as a fallback.

5. **Output:**
   - The function returns a `Range` object, which can be used in Excel formulas to reference cells on the previous sheet (or the current sheet if it’s the first one).

#### ** Example Usage: **
Imagine a workbook with three sheets: `Sheet1`, `Sheet2`, and `Sheet3`, in that order.

- **Scenario 1: Calling from Sheet2**
  - Formula in `Sheet2!B1`: `=PrevSheet(A1)`
  - `A1` on `Sheet2` is the input range (`rng`).
  - `ws` is `Sheet2`, with `Index = 2`.
  - Since `2 > 1`, it targets `Sheet1` (`Index = 1`).
  - Returns `Sheet1!A1`.
  - If `Sheet1!A1` contains `10`, the formula in `Sheet2!B1` evaluates to `10`.

- **Scenario 2: Calling from Sheet1**
  - Formula in `Sheet1!B1`: `=PrevSheet(A1)`
  - `ws` is `Sheet1`, with `Index = 1`.
  - Since `1 > 1` is false, it falls back to `Sheet1`.
  - Returns `Sheet1!A1`.
  - If `Sheet1!A1` contains `5`, the formula in `Sheet1!B1` evaluates to `5`.

- **Scenario 3: Multi-Cell Range**
  - Formula in `Sheet3!C1`: `=SUM(PrevSheet(A1:A3))`
  - `rng` is `Sheet3!A1:A3`.
  - `ws` is `Sheet3`, with `Index = 3`.
  - Targets `Sheet2!A1:A3`.
  - Sums the values in `Sheet2!A1:A3`.

#### ** Practical Use Cases: **
- **Cross-Sheet Calculations:** Useful for comparing or aggregating data across sequential sheets (e.g., monthly reports where each sheet is a month).
- **Dynamic References:** If sheets are added or reordered, the function adapts by always pointing to the "previous" sheet based on the current index.
- **Fallback Safety:** The `Else` clause ensures the function doesn’t error out on the first sheet, making it robust for edge cases.

---

#### ** Key Features: **
- **Dynamic:** Adjusts to the workbook’s sheet order at runtime.
- **Simple:** Focuses solely on referencing the previous sheet’s range without additional logic.
- **Range-Based Output:** Returns a `Range` object, so it can be combined with other Excel functions like `SUM`, `AVERAGE`, etc.

----------------------------------------------------------------------------------------------
----------------------------------------------------------------------------------------------

## Custom Actions

### ** AddSheetChangeAction ** - * Add to ThisWorkbook *

#### Overview of the Code
This is a class module named `ThisWorkbook`, which is a special object in Excel VBA representing the workbook itself. Code placed here can respond to workbook-level events, such as opening, closing, or sheet-specific actions.
This `ThisWorkbook` module ensures that all formulas are recalculated whenever you switch sheets, edit cells, or leave a sheet. It’s a heavy-handed approach to keep everything up-to-date, likely paired with volatile functions like `PrevSheet` to guarantee dynamic references work correctly. However, it sacrifices performance for certainty, so it’s best suited for smaller workbooks or scenarios where absolute consistency outweighs speed.

---
**Code:**

Private Sub Workbook_SheetActivate(ByVal Sh As Object)
    Application.CalculateFull ' Recalculate all formulas when switching sheets
End Sub

Private Sub Workbook_SheetChange(ByVal Sh As Object, ByVal Target As Range)
    Application.CalculateFull ' Recalculate all formulas when any cell changes
End Sub

Private Sub Workbook_SheetDeactivate(ByVal Sh As Object)
    Application.CalculateFull ' Recalculate all formulas when leaving a sheet
End Sub

---

#### Breakdown of Components

##### 1. **Class Header and Attributes**
- `VERSION 1.0 CLASS`: Indicates this is a class module (standard for `ThisWorkbook`).
- `MultiUse = -1 'True`: Allows multiple instances (though not directly relevant here since `ThisWorkbook` is a singleton object).
- Attributes:
  - `VB_Name = "ThisWorkbook"`: Names the module.
  - `VB_GlobalNameSpace = False`: Not accessible globally outside the project.
  - `VB_Creatable = False`: Cannot be instantiated manually (it’s predefined).
  - `VB_PredeclaredId = True`: The object exists by default (Excel creates it).
  - `VB_Exposed = True`: Exposed to other VBA projects or macros.

These attributes are standard for the `ThisWorkbook` object and define its behavior in the VBA environment.

##### 2. **Event Handlers**
The code includes three private subroutines, each tied to a specific workbook event. All three use the same action: `Application.CalculateFull`.

- **`Workbook_SheetActivate(ByVal Sh As Object)`**
  - **Trigger:** Fires when a sheet in the workbook is activated (e.g., you click its tab to switch to it).
  - **Parameter:** `Sh` is the sheet object being activated (e.g., `Sheet1`).
  - **Action:** `Application.CalculateFull` forces a full recalculation of all formulas in the entire workbook, regardless of whether they need it.

- **`Workbook_SheetChange(ByVal Sh As Object, ByVal Target As Range)`**
  - **Trigger:** Fires when any cell on any sheet in the workbook is changed (e.g., you type a value or edit a cell).
  - **Parameters:**
    - `Sh`: The sheet where the change occurred.
    - `Target`: The range of cells that were modified.
  - **Action:** `Application.CalculateFull` recalculates all formulas in the workbook.

- **`Workbook_SheetDeactivate(ByVal Sh As Object)`**
  - **Trigger:** Fires when a sheet is deactivated (e.g., you switch away from it to another sheet).
  - **Parameter:** `Sh` is the sheet being deactivated.
  - **Action:** `Application.CalculateFull` recalculates all formulas in the workbook.

#### What `Application.CalculateFull` Does
- **Purpose:** Forces Excel to recalculate every formula in every open workbook, even if Excel’s calculation engine doesn’t think it’s necessary (e.g., no dependencies changed).
- **Contrast with Alternatives:**
  - `Application.Calculate`: Recalculates only what Excel deems necessary based on dependencies.
  - `Worksheet.Calculate`: Recalculates only the specified worksheet.
- **Impact:** Ensures all volatile functions (like the `PrevSheet` function from your earlier question) and any potentially stale calculations are updated.

#### How It Works in Practice
- **Sheet Activation:** Switch from `Sheet1` to `Sheet2` → `Workbook_SheetActivate` runs for `Sheet2`, recalculating everything.
- **Sheet Change:** Edit `Sheet1!A1` → `Workbook_SheetChange` runs for `Sheet1` and the changed cell (`A1`), recalculating everything.
- **Sheet Deactivation:** Switch from `Sheet1` to `Sheet2` → `Workbook_SheetDeactivate` runs for `Sheet1`, recalculating everything.

This creates a highly aggressive recalculation strategy: almost any interaction with the workbook triggers a full recalculation.

#### Example Scenario
- **Workbook Setup:** Three sheets (`Sheet1`, `Sheet2`, `Sheet3`) with formulas, including volatile ones like `=PrevSheet(A1)`.
- **Actions:**
  1. Click `Sheet2` → Full recalculation.
  2. Type `5` in `Sheet2!B1` → Full recalculation.
  3. Switch to `Sheet3` → Full recalculation when leaving `Sheet2`, then another when activating `Sheet3`.
- **Result:** Every action ensures all formulas reflect the latest state, especially useful if you’re using custom functions like `PrevSheet` that rely on sheet order or dynamic data.

#### Practical Use Cases
- **Ensuring Consistency:** If your workbook uses volatile custom functions (e.g., `PrevSheet`, `SumByColor`) or external data links that Excel might not automatically update, this forces everything to stay current.
- **Debugging or Testing:** Useful during development to ensure formulas behave as expected after every change or navigation.
- **Complex Dependencies:** In workbooks where sheet order or inter-sheet references (like `PrevSheet`) matter, this guarantees no stale values persist.

#### Key Implications
- **Performance:** `CalculateFull` can be slow in large workbooks with many formulas, sheets, or complex calculations. Triggering it on every sheet activation, change, or deactivation could make the workbook feel sluggish.
- **Redundancy:** Recalculating on *every* event might be overkill. For example, a change on `Sheet1` doesn’t necessarily require recalculating `Sheet2` unless there’s a dependency.
- **User Experience:** Frequent recalculations might interrupt workflow, especially if the workbook is shared or used interactively.

#### Potential Issues
- **Lag:** In a workbook with thousands of formulas or large datasets, users might notice delays after every action.
- **Unnecessary Recalcs:** If Excel’s automatic calculation mode is already on (`Application.Calculation = xlCalculationAutomatic`), this could duplicate effort, as Excel already recalculates dependencies.
- **No Control:** There’s no condition to limit recalculation (e.g., only for specific sheets or changes), making it a blunt tool.

#### Possible Enhancements
- **Selective Recalculation:** Add logic to only recalculate specific sheets:
  ```vba
  Private Sub Workbook_SheetChange(ByVal Sh As Object, ByVal Target As Range)
      Sh.Calculate ' Recalculate only the changed sheet
  End Sub
  ```
- **Toggle Control:** Use a global variable or cell value to enable/disable forced recalculation:
  ```vba
  Private Sub Workbook_SheetActivate(ByVal Sh As Object)
      If ThisWorkbook.Sheets("Settings").Range("A1").Value = True Then
          Application.CalculateFull
      End If
  End Sub
  ```
- **Debounce:** Add a delay or check to avoid multiple recalculations in quick succession (though this is trickier in VBA).

----------------------------------------------------------------------------------------------

### `CopyConditionalFillColor` - *Add as Module and UserForm*

#### **Purpose:**
The `CopyConditionalFillColor` macro copies displayed formatting properties (e.g., fill colors, font colors, bold, italic, borders, number formats) from a source range to a target range in an Excel worksheet. It uses the `DisplayFormat` property to capture formatting applied via **conditional formatting**, which is not accessible through standard copy operations. A UserForm interface allows users to select source and target worksheets, specify source and target ranges with mouse-based selection, and choose properties to copy, with validation to ensure correct inputs.

This macro enables you to convert conditional formatting into static formats, allowing you to remove conditional formatting rules while preserving the displayed formatting at the time of copying. It’s ideal for “locking in” formatting for sharing or further processing.

#### **How to Install:**
1. **Add the Module and UserForm:**
   - Place `CopyConditionalFillColorForm.frx` in the same directory as `CopyConditionalFillColorForm.frm` before importing.
   - Import **BOTH** `CopyConditionalFillColor.bas` and `CopyConditionalFillColorForm.frm` into your VBA project (see [Steps to Install Custom VBA Formulas via Import](#steps-to-install-custom-vba-formulas-via-import--bas--or-cls-)).

2. ##### **Manual Creation (Fallback):**
   - See [Manual Form Creation Instructions](docs/CopyConditionalFillColorForm - Manual Form Creation Instructions.md)

#### **How It Works:**

- **Inputs (via UserForm):**
  - **Source Worksheet:** Selected from a dropdown listing all worksheets in the workbook.
  - **Target Worksheet:** Selected from a dropdown listing all worksheets in the workbook.
  - **Source Range:** Specified via a TextBox, with a “Select” button for mouse-based range selection.
  - **Target Range:** Specified via a TextBox, with a “Select” button for mouse-based range selection.
  - **Properties to Copy:** Chosen via CheckBoxes for Fill Color, Font Color, Bold, Italic, Borders, and Number Format.

- **Logic:**
  1. Displays a UserForm (`CopyConditionalFillColorForm`) with:
     - ComboBoxes for selecting source and target worksheets (default: active sheet).
     - TextBoxes for source and target ranges (default: `A1`).
     - “Select” buttons to open a range selection dialog (`Application.InputBox` with `Type:=8`), ensuring ranges are on the selected worksheets.
     - CheckBoxes to choose formatting properties (Fill Color and Font Color checked by default).
     - OK and Cancel buttons.

  2. When the user clicks OK:
     - Validates inputs:
       - Ensures source and target worksheets are valid.
       - Verifies source and target ranges are valid.
       - Checks that source and target ranges have the same number of cells.
       - Confirms at least one property is selected.
     - If validation fails, displays an error message and keeps the UserForm open for corrections.
     - If valid, retrieves the source range’s displayed formatting using `DisplayFormat` properties (e.g., `DisplayFormat.Interior.Color`, `DisplayFormat.Font.Color`).
     - Applies the selected properties as static formats to the target range, adjusting for relative cell positions.

  3. If the user clicks Cancel, the macro exits without changes.

  4. Displays a success message when complete.
- **Key Features:**
  - Captures conditional formatting appearance using `DisplayFormat`, not just manual formats.
  - Supports multiple formatting properties (not limited to fill colors).
  - Mouse-based range selection mimics Excel’s **Insert Function** dialog for intuitive input.
  - Validates range sizes and inputs to prevent errors.
  - Flexible source and target worksheet selection via dropdowns.

#### **Example Use Case:**
- **Setup:** On `Sheet1`, range `A1:C10` has conditional formatting (e.g., values > 10: yellow fill, red font, bold; values < 5: blue border, italic, currency format). You want to copy these formats to `Sheet2!D1:F10` as static formats.
- **Action:**
  - Run the macro, select `Sheet1` in the source worksheet dropdown and `Sheet2` in the target worksheet dropdown.
  - Click “Select” for Source Range, choose `A1:C10` with the mouse.
  - Click “Select” for Target Range, choose `D1:F10` on `Sheet2`.
  - Check desired properties (e.g., Fill Color, Font Color, Bold).
  - Click OK.
- **Result:** `Sheet2!D1:F10` has static formats matching the displayed formatting of `Sheet1!A1:C10` (e.g., yellow fills, red fonts, bold where applicable).


#### **How to Use:**

1. **Add the Module and UserForm:** (see [Steps to Install Custom VBA Formulas via Import](#steps-to-install-custom-vba-formulas-via-import--bas--or-cls-)).

2. **Run the Macro:**
   - Go to `Developer > Macros` (or `Alt + F8`), select `CopyConditionalFillColor`, and click `Run`.
   - Alternatively, assign the macro to a button or run it via the VBA Editor.

3. **Use the UserForm:**
   - **Select Worksheets:** Choose the source and target worksheets from the dropdowns (default: active sheet).
   - **Select Source Range:** Click the “Select” button next to “Source Range” to choose a range with the mouse (e.g., `A1:C10`) on the source worksheet. Alternatively, type the range address.
   - **Select Target Range:** Click the “Select” button next to “Target Range” to choose a range (e.g., `D1:F10`) on the target worksheet. Alternatively, type the address.
   - **Choose Properties:** Check the properties to copy (e.g., Fill Color, Font Color). At least one must be selected.
   - **Confirm:** Click OK to apply the formats. If inputs are invalid (e.g., mismatched range sizes, invalid worksheets), an error message will prompt you to correct them. Click Cancel to exit without changes.

4. **Verify Output:**
   - Check the target range for the applied static formats (visible in `Home > Format Cells > Fill`, `Font`, etc.).
   - The source range’s conditional formatting remains unchanged.

5. **Post-Processing (Optional):**
   - **Delete Conditional Formatting Rules:** If no longer needed, remove rules from the source or target range (`Home > Conditional Formatting > Clear Rules`).
   - **Remove Macro:** If the macro is no longer needed, delete the module and UserForm, then save the workbook as `.xlsx` for broader compatibility.


#### **Limitations:**
- **Static Formats:** Copied formats are static, losing any conditional formatting rules in the target range.
- **Performance:** Processing large ranges (e.g., thousands of cells) may be slow due to cell-by-cell copying.
- **UserForm Dependency:** Requires importing both `.frm` and `.frx` files correctly for the UserForm to function.
- **Manual Input Option:** Users can type range addresses, which may lead to errors if invalid (e.g., `Z1:AA10`), though validation catches most issues.


#### **Troubleshooting:**

- **Import Errors:**
  - Ensure `CopyConditionalFillColorForm.frx` is in the same directory as `CopyConditionalFillColorForm.frm`.
  - Use Notepad to verify the `.frm` file starts with `VERSION 5.00` and includes `OleObjectBlob = "CopyConditionalFillColorForm.frx":0000`.
  - If errors persist, manually create or update the UserForm in the VBE and export new `.frm` and `.frx` files (see [Manual Creation](#manual-creation)).

- **Range Selection Issues:**
  - Ensure macros are enabled (`Developer > Macro Security > Enable All Macros` for testing).
  - If the range selection dialog fails, test `Application.InputBox(Type:=8)` in a separate macro.
  - Confirm selected ranges are on the correct worksheets (e.g., source range on the source worksheet).

- **Validation Errors:**
  - If you see “Source and target ranges must have the same number of cells,” ensure the selected ranges match in size (e.g., both 3x3 cells).
  - Check that worksheet and range inputs are valid (e.g., no typos in range addresses, valid worksheet names).
  - Ensure at least one property CheckBox is selected.

- **Formatting Not Applied:**
  - Verify the source range has conditional formatting applied (`Home > Conditional Formatting > Manage Rules`).
  - Check that the correct worksheets are selected in the dropdowns.

- **Excel Version:** Some features (e.g., `DisplayFormat`) require Excel 2010 or later. Test on your version (e.g., 2016, 365).
  
  -----------------------------------------------------------------------------------------------
