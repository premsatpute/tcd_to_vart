hey there , 
This application provides a convenient way to convert Test Case Design (TCD) Excel files into a structured VART (Verification and Automation Ready Test) format. It extracts relevant information, organizes test steps, and generates an Excel file suitable for automation or further analysis.

Features:
TCD Data Preprocessing: Loads TCD Excel files (.xlsx), selects essential columns (Labels, Action, Expected Results), and performs data cleaning and normalization.
Intelligent Step Extraction: Parses "Action" and "Expected Results" fields to extract individual test steps. It intelligently handles numbered steps and includes a specific logic for "battery reconnect" sequences, expanding them into a series of relay operations.
VART Sheet Generation: Organizes the extracted test cases into a VART-style Excel sheet.
Categorization: Groups test cases by Test_Case_Type (e.g., Logical Combination, Failure Modes, Power Modes, Configuration, Voltage Modes).
Feature-based Structure: Arranges test cases under their respective features, providing a clear hierarchy.
Keyword-Value Pairs: Transforms test steps into keyword-value pairs, which is a common format for automation frameworks.
"END" Keyword: Automatically appends an "END" keyword with a "yes" value at the conclusion of each test script, signaling the end of a test sequence.
Excel Formatting: Applies distinct background colors to differentiate features, categories, and keyword headers, improving readability.

TCD File Requirements:
Your input TCD Excel file should contain at least the following columns:
Labels: Contains test case labels, from which "Test_Case_Type" and "Sub_Feature" are derived.
Action: Describes the actions to be performed for the test case. Steps should ideally be numbered or clearly delineated.
Expected Results: Describes the expected outcomes for the test case.
